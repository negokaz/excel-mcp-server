
FROM node:20-slim AS release

# Set the working directory
WORKDIR /app

RUN npm install -g @negokaz/excel-mcp-server@0.11.1

# Command to run the application
ENTRYPOINT ["excel-mcp-server"]
