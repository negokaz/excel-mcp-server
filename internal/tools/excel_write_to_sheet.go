package tools

import (
	"context"
	"fmt"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
	"github.com/xuri/excelize/v2"
)

type ExcelWriteToSheetArguments struct {
	FileAbsolutePath string     `zog:"fileAbsolutePath"`
	SheetName        string     `zog:"sheetName"`
	NewSheet         bool       `zog:"newSheet"`
	Range            string     `zog:"range"`
	Values           [][]string `zog:"values"`
}

var excelWriteToSheetArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"newSheet":         z.Bool().Required().Default(false),
	"range":            z.String().Required(),
	"values":           z.Slice(z.Slice(z.String())).Required(),
})

func AddExcelWriteToSheetTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_write_to_sheet",
		mcp.WithDescription("Write values to the Excel sheet"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithBoolean("newSheet",
			mcp.Required(),
			mcp.Description("Create a new sheet if true, otherwise write to the existing sheet"),
		),
		mcp.WithString("range",
			mcp.Required(),
			mcp.Description("Range of cells in the Excel sheet (e.g., \"A1:C10\")"),
		),
		mcp.WithArray("values",
			mcp.Required(),
			mcp.Description("Values to write to the Excel sheet. If the value is a formula, it should start with \"=\""),
			mcp.Items(map[string]any{
				"type": "array",
				"items": map[string]any{
					"anyOf": []any{
						map[string]any{
							"type": "string",
						},
						map[string]any{
							"type": "number",
						},
						map[string]any{
							"type": "boolean",
						},
						map[string]any{
							"type": "null",
						},
					},
				},
			}),
		),
	), handleWriteToSheet)
}

func handleWriteToSheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelWriteToSheetArguments{}
	issues := excelWriteToSheetArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}

	// zog が any type のスキーマをサポートしていないため、自力で実装
	valuesArg, ok := request.GetArguments()["values"].([]any)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("values must be a 2D array"), nil
	}
	values := make([][]any, len(valuesArg))
	for i, v := range valuesArg {
		value, ok := v.([]any)
		if !ok {
			return imcp.NewToolResultInvalidArgumentError("values must be a 2D array"), nil
		}
		values[i] = value
	}

	return writeSheet(args.FileAbsolutePath, args.SheetName, args.NewSheet, args.Range, values)
}

func writeSheet(fileAbsolutePath string, sheetName string, newSheet bool, rangeStr string, values [][]any) (*mcp.CallToolResult, error) {
	workbook, closeFn, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer closeFn()

	if newSheet {
		if err := workbook.CreateNewSheet(sheetName); err != nil {
			return nil, err
		}
	}

	// シートの取得
	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	defer worksheet.Release()

	startCol, startRow, endCol, endRow, err := excel.ParseRange(rangeStr)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}

	// データの整合性チェック
	rangeRowSize := endRow - startRow + 1
	if len(values) != rangeRowSize {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("number of rows in data (%d) does not match range size (%d)", len(values), rangeRowSize)), nil
	}

	// データの書き込み
	wroteFormula := false
	for i, row := range values {
		rangeColumnSize := endCol - startCol + 1
		if len(row) != rangeColumnSize {
			return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("number of columns in row %d (%d) does not match range size (%d)", i, len(row), rangeColumnSize)), nil
		}
		for j, cellValue := range row {
			cell, err := excelize.CoordinatesToCellName(startCol+j, startRow+i)
			if err != nil {
				return nil, err
			}
			if cellStr, ok := cellValue.(string); ok && isFormula(cellStr) {
				// if cellValue is formula, set it as formula
				err = worksheet.SetFormula(cell, cellStr)
				wroteFormula = true
			} else {
				// if cellValue is not formula, set it as value
				err = worksheet.SetValue(cell, cellValue)
			}
			if err != nil {
				return nil, err
			}
		}
	}

	if err := workbook.Save(); err != nil {
		return nil, err
	}

	// HTMLテーブルの生成
	var table *string
	if wroteFormula {
		table, err = CreateHTMLTableOfFormula(worksheet, startCol, startRow, endCol, endRow)
	} else {
		table, err = CreateHTMLTableOfValues(worksheet, startCol, startRow, endCol, endRow)
	}
	if err != nil {
		return nil, err
	}
	html := "<h2>Written Sheet</h2>\n"
	html += *table + "\n"
	html += "<h2>Metadata</h2>\n"
	html += "<ul>\n"
	html += fmt.Sprintf("<li>backend: %s</li>\n", workbook.GetBackendName())
	html += fmt.Sprintf("<li>sheet name: %s</li>\n", sheetName)
	html += fmt.Sprintf("<li>read range: %s</li>\n", rangeStr)
	html += "</ul>\n"
	html += "<h2>Notice</h2>\n"
	html += "<p>Values wrote successfully.</p>\n"

	return mcp.NewToolResultText(html), nil
}

func isFormula(value string) bool {
	return len(value) > 0 && value[0] == '='
}
