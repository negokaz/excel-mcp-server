package tools

import (
	"fmt"

	"github.com/devlights/goxcel"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/xuri/excelize/v2"
)

// PagingStrategy はページング範囲の計算戦略を定義するインターフェース
type PagingStrategy interface {
	// CalculatePagingRanges は利用可能なページング範囲のリストを返す
	CalculatePagingRanges() []string
	// ValidatePagingRange は指定された範囲が有効かどうかを検証する
	ValidatePagingRange(rangeStr string) error
}

// FixedSizePagingStrategy は固定サイズでページング範囲を計算する戦略
type FixedSizePagingStrategy struct {
	pageSize  int
	workbook  *excelize.File
	sheetName string
	dimension string
}

// NewFixedSizePagingStrategy は新しいFixedSizePagingStrategyインスタンスを生成する
func NewFixedSizePagingStrategy(pageSize int, workbook *excelize.File, sheetName string) (*FixedSizePagingStrategy, error) {
	if pageSize <= 0 {
		pageSize = 5000 // デフォルト値
	}

	// シート名の検証
	index, err := workbook.GetSheetIndex(sheetName)
	if err != nil || index == -1 {
		return nil, fmt.Errorf("sheet %s not found", sheetName)
	}

	// まずシートの次元情報を取得
	dimension, err := workbook.GetSheetDimension(sheetName)
	if err != nil {
		return nil, err
	}

	// 次元情報から開始位置を取得
	startCol, startRow, endCol, endRow, err := ParseRange(dimension)
	if err != nil {
		return nil, err
	}

	// データを含むセルを探すために、シートをスキャン
	maxRow := endRow
	maxCol := endCol

	// 行方向のスキャン（最大1000行まで）
	for row := endRow + 1; row <= 1000; row++ {
		hasData := false
		for col := startCol; col <= endCol; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				continue
			}
			value, _ := workbook.GetCellValue(sheetName, cell)
			if value != "" {
				hasData = true
				maxRow = row
				break
			}
		}
		if !hasData {
			break
		}
	}

	// 列方向のスキャン
	for col := endCol + 1; col <= 26; col++ { // A-Z の範囲
		hasData := false
		for row := startRow; row <= maxRow; row++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				continue
			}
			value, _ := workbook.GetCellValue(sheetName, cell)
			if value != "" {
				hasData = true
				maxCol = col
				break
			}
		}
		if !hasData {
			break
		}
	}

	// 新しい次元情報を作成
	startCell, _ := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, _ := excelize.CoordinatesToCellName(maxCol, maxRow)
	newDimension := fmt.Sprintf("%s:%s", startCell, endCell)

	return &FixedSizePagingStrategy{
		pageSize:  pageSize,
		workbook:  workbook,
		sheetName: sheetName,
		dimension: newDimension,
	}, nil
}

// CalculatePagingRanges は固定サイズに基づいてページング範囲のリストを生成する
func (s *FixedSizePagingStrategy) CalculatePagingRanges() []string {
	startCol, startRow, endCol, endRow, err := ParseRange(s.dimension)
	if err != nil {
		return []string{}
	}

	totalCols := endCol - startCol + 1
	cellsPerPage := s.pageSize

	// 1ページあたりの行数を計算
	rowsPerPage := cellsPerPage / totalCols
	if rowsPerPage < 1 {
		rowsPerPage = 1
	}

	var ranges []string
	currentRow := startRow
	for currentRow <= endRow {
		pageEndRow := currentRow + rowsPerPage - 1
		if pageEndRow > endRow {
			pageEndRow = endRow
		}

		startRange, _ := excelize.CoordinatesToCellName(startCol, currentRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, pageEndRow)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))

		currentRow = pageEndRow + 1
	}

	return ranges
}

// ValidatePagingRange は指定された範囲が有効かどうかを検証する
func (s *FixedSizePagingStrategy) ValidatePagingRange(rangeStr string) error {
	startCol, startRow, endCol, endRow, err := ParseRange(rangeStr)
	if err != nil {
		return fmt.Errorf("invalid range format: %v", err)
	}

	dimStartCol, dimStartRow, dimEndCol, dimEndRow, err := ParseRange(s.dimension)
	if err != nil {
		return fmt.Errorf("invalid dimension format: %v", err)
	}

	// 範囲がシートの次元内に収まっているか確認
	if startCol < dimStartCol || startRow < dimStartRow ||
		endCol > dimEndCol || endRow > dimEndRow {
		return fmt.Errorf("range %s is outside sheet dimensions %s",
			rangeStr, s.dimension)
	}

	// セル数が pageSize を超えていないか確認
	cellCount := (endRow - startRow + 1) * (endCol - startCol + 1)
	if cellCount > s.pageSize {
		return fmt.Errorf("range contains %d cells, exceeding page size of %d",
			cellCount, s.pageSize)
	}

	return nil
}

func NewGoxcelPagingStrategy(pageSize int, worksheet *goxcel.Worksheet) (PagingStrategy, error) {
	if worksheet == nil {
		return nil, fmt.Errorf("worksheet is nil")
	}

	printAreaPagingStrategy, err := NewPrintAreaPagingStrategy(worksheet)
	if err != nil {
		return nil, err
	}
	printArea, err := printAreaPagingStrategy.getPrintArea()
	if err != nil {
		return nil, err
	}
	if printArea == "" {
		return NewGoxcelFixedSizePagingStrategy(pageSize, worksheet)
	} else {
		return printAreaPagingStrategy, nil
	}
}

// GoxcelFixedSizePagingStrategy は Goxcel を使用した固定サイズでページング範囲を計算する戦略
type GoxcelFixedSizePagingStrategy struct {
	pageSize  int
	worksheet *goxcel.Worksheet
	dimension string
}

// NewGoxcelFixedSizePagingStrategy は新しい GoxcelFixedSizePagingStrategy インスタンスを生成する
func NewGoxcelFixedSizePagingStrategy(pageSize int, worksheet *goxcel.Worksheet) (*GoxcelFixedSizePagingStrategy, error) {
	if pageSize <= 0 {
		pageSize = 5000 // デフォルト値
	}

	if worksheet == nil {
		return nil, fmt.Errorf("worksheet is nil")
	}

	// UsedRange を使用して使用範囲を取得
	usedRange, err := worksheet.UsedRange()
	if err != nil {
		return nil, fmt.Errorf("failed to get used range: %w", err)
	}

	// 範囲の情報を取得
	address, err := FetchRangeAddress(usedRange)
	if err != nil {
		return nil, fmt.Errorf("failed to get range address: %w", err)
	}

	return &GoxcelFixedSizePagingStrategy{
		pageSize:  pageSize,
		worksheet: worksheet,
		dimension: address,
	}, nil
}

// CalculatePagingRanges は固定サイズに基づいてページング範囲のリストを生成する
func (s *GoxcelFixedSizePagingStrategy) CalculatePagingRanges() []string {
	startCol, startRow, endCol, endRow, err := ParseRange(s.dimension)
	if err != nil {
		return []string{}
	}

	totalCols := endCol - startCol + 1
	cellsPerPage := s.pageSize

	// 1ページあたりの行数を計算
	rowsPerPage := cellsPerPage / totalCols
	if rowsPerPage < 1 {
		rowsPerPage = 1
	}

	var ranges []string
	currentRow := startRow
	for currentRow <= endRow {
		pageEndRow := currentRow + rowsPerPage - 1
		if pageEndRow > endRow {
			pageEndRow = endRow
		}

		startRange, _ := excelize.CoordinatesToCellName(startCol, currentRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, pageEndRow)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))

		currentRow = pageEndRow + 1
	}

	return ranges
}

// ValidatePagingRange は指定された範囲が有効かどうかを検証する
func (s *GoxcelFixedSizePagingStrategy) ValidatePagingRange(rangeStr string) error {
	startCol, startRow, endCol, endRow, err := ParseRange(rangeStr)
	if err != nil {
		return fmt.Errorf("invalid range format: %v", err)
	}

	dimStartCol, dimStartRow, dimEndCol, dimEndRow, err := ParseRange(s.dimension)
	if err != nil {
		return fmt.Errorf("invalid dimension format: %v", err)
	}

	// 範囲がシートの次元内に収まっているか確認
	if startCol < dimStartCol || startRow < dimStartRow ||
		endCol > dimEndCol || endRow > dimEndRow {
		return fmt.Errorf("range %s is outside sheet dimensions %s",
			rangeStr, s.dimension)
	}

	// セル数が pageSize を超えていないか確認
	cellCount := (endRow - startRow + 1) * (endCol - startCol + 1)
	if cellCount > s.pageSize {
		return fmt.Errorf("range contains %d cells, exceeding page size of %d",
			cellCount, s.pageSize)
	}

	return nil
}

// PrintAreaPagingStrategy は印刷範囲とページ区切りに基づいてページング範囲を計算する戦略
type PrintAreaPagingStrategy struct {
	worksheet *goxcel.Worksheet
}

// NewPrintAreaPagingStrategy は新しいPrintAreaPagingStrategyインスタンスを生成する
func NewPrintAreaPagingStrategy(worksheet *goxcel.Worksheet) (*PrintAreaPagingStrategy, error) {
	if worksheet == nil {
		return nil, fmt.Errorf("worksheet is nil")
	}
	return &PrintAreaPagingStrategy{
		worksheet: worksheet,
	}, nil
}

// getPrintArea は印刷範囲を取得する
func (s *PrintAreaPagingStrategy) getPrintArea() (string, error) {
	ps, err := s.worksheet.PageSetup()
	if err != nil {
		return "", fmt.Errorf("failed to get PageSetup: %w", err)
	}
	area, err := oleutil.GetProperty(ps.ComObject(), "PrintArea")
	if err != nil {
		return "", fmt.Errorf("failed to get PrintArea: %w", err)
	}
	address := area.ToString()
	if address == "" {
		return "", nil
	} else {
		return address, nil
	}
}

// getHPageBreaksPositions はページ区切りの位置情報を取得する
func (s *PrintAreaPagingStrategy) getHPageBreaksPositions() ([]int, error) {
	hpb, err := s.worksheet.HPageBreaks()
	if err != nil {
		return nil, fmt.Errorf("failed to get HPageBreaks: %w", err)
	}
	count, err := hpb.Count()
	if err != nil {
		return nil, fmt.Errorf("failed to get HPageBreaks count: %w", err)
	}

	positions := make([]int, 0, int(count))
	for i := 1; i <= int(count); i++ {
		pageBreak, err := hpb.Item(i)
		if err != nil {
			return nil, fmt.Errorf("failed to get HPageBreak item %d: %w", i, err)
		}
		location, err := pageBreak.Location()
		if err != nil {
			return nil, fmt.Errorf("failed to get HPageBreak row %d: %w", i, err)
		}
		row, err := location.Row()
		if err != nil {
			return nil, fmt.Errorf("failed to get HPageBreak row %d: %w", i, err)
		}
		positions = append(positions, int(row))
	}

	return positions, nil
}

// calculateRangesFromBreaks は印刷範囲とページ区切りから範囲のリストを生成する
func (s *PrintAreaPagingStrategy) calculateRangesFromBreaks(printArea string, breaks []int) []string {
	if printArea == "" {
		return []string{}
	}

	startCol, startRow, endCol, endRow, err := ParseRange(printArea)
	if err != nil {
		return []string{}
	}

	ranges := make([]string, 0)
	currentRow := startRow

	// ページ区切りがない場合は印刷範囲全体を1つの範囲として扱う
	if len(breaks) == 0 {
		startRange, _ := excelize.CoordinatesToCellName(startCol, startRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, endRow)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))
		return ranges
	}

	// ページ区切りで範囲を分割
	for _, breakRow := range breaks {
		if breakRow <= startRow || breakRow > endRow {
			continue
		}

		startRange, _ := excelize.CoordinatesToCellName(startCol, currentRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, breakRow-1)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))

		currentRow = breakRow
	}

	// 最後の範囲を追加
	if currentRow <= endRow {
		startRange, _ := excelize.CoordinatesToCellName(startCol, currentRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, endRow)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))
	}

	return ranges
}

// CalculatePagingRanges は印刷範囲とページ区切りに基づいてページング範囲のリストを生成する
func (s *PrintAreaPagingStrategy) CalculatePagingRanges() []string {
	printArea, err := s.getPrintArea()
	if err != nil {
		return []string{}
	}

	breaks, err := s.getHPageBreaksPositions()
	if err != nil {
		return []string{}
	}

	return s.calculateRangesFromBreaks(printArea, breaks)
}

// ValidatePagingRange は指定された範囲が印刷範囲内に収まっているか検証する
func (s *PrintAreaPagingStrategy) ValidatePagingRange(rangeStr string) error {
	printArea, err := s.getPrintArea()
	if err != nil {
		return fmt.Errorf("failed to get print area: %w", err)
	}
	if printArea == "" {
		return fmt.Errorf("print area is not set")
	}

	// 印刷範囲を解析
	paStartCol, paStartRow, paEndCol, paEndRow, err := ParseRange(printArea)
	if err != nil {
		return fmt.Errorf("invalid print area format: %w", err)
	}

	// 検証対象の範囲を解析
	startCol, startRow, endCol, endRow, err := ParseRange(rangeStr)
	if err != nil {
		return fmt.Errorf("invalid range format: %w", err)
	}

	// 範囲が印刷範囲内に収まっているか確認
	if startCol < paStartCol || startRow < paStartRow ||
		endCol > paEndCol || endRow > paEndRow {
		return fmt.Errorf("range %s is outside print area %s",
			rangeStr, printArea)
	}

	// ページ区切り位置と一致するか確認
	ranges := s.CalculatePagingRanges()
	for _, r := range ranges {
		if r == rangeStr {
			return nil
		}
	}

	return fmt.Errorf("range %s does not match any page break boundary", rangeStr)
}

// PagingRangeService はページング処理を提供するサービス
type PagingRangeService struct {
	strategy PagingStrategy
}

// NewPagingRangeService は新しいPagingRangeServiceインスタンスを生成する
func NewPagingRangeService(strategy PagingStrategy) *PagingRangeService {
	return &PagingRangeService{strategy: strategy}
}

// GetPagingRanges は利用可能なページング範囲のリストを返す
func (s *PagingRangeService) GetPagingRanges() []string {
	return s.strategy.CalculatePagingRanges()
}

// ValidatePagingRange は指定された範囲が有効かどうかを検証する
func (s *PagingRangeService) ValidatePagingRange(rangeStr string) error {
	return s.strategy.ValidatePagingRange(rangeStr)
}

// FilterRemainingPagingRanges は未読の範囲のリストを返す
func (s *PagingRangeService) FilterRemainingPagingRanges(allRanges []string, knownRanges []string) []string {
	if len(knownRanges) == 0 {
		return allRanges
	}

	// knownRanges をマップに変換
	knownMap := make(map[string]bool)
	for _, r := range knownRanges {
		knownMap[r] = true
	}

	// 未読の範囲を抽出
	remaining := make([]string, 0)
	for _, r := range allRanges {
		if !knownMap[r] {
			remaining = append(remaining, r)
		}
	}

	return remaining
}
