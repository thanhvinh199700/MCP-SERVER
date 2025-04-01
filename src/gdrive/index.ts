#!/usr/bin/env node

import { authenticate } from "@google-cloud/local-auth";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListResourcesRequestSchema,
  ListToolsRequestSchema,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import fs from "fs";
import { google, sheets_v4 } from 'googleapis';
import path from "path";
import { fileURLToPath } from 'url';

// Định nghĩa các interface cho Google Sheets
interface SheetProperties {
  title?: string;
  sheetId?: number;
  gridProperties?: {
    rowCount?: number;
    columnCount?: number;
  };
}

interface Sheet {
  properties?: SheetProperties;
}

interface SpreadsheetResponse {
  data: {
    sheets?: Sheet[];
  };
}

interface CellData {
  formattedValue?: string;
}

interface RowData {
  values?: CellData[];
}

interface SheetData {
  rowData?: RowData[];
}

interface GridData {
  data?: SheetData[];
}

// Thêm interface cho Google Sheets Values Response
interface SheetsValueRange {
  values?: any[][];
  majorDimension?: string;
  range?: string;
}

interface SheetsResponse {
  data: SheetsValueRange;
}

// Thêm interface cho chỉnh sửa Google Sheets
interface UpdateSheetRequest {
  fileId: string;
  sheetName?: string;
  range: string;
  values: any[][];
}

const drive = google.drive("v3");
const sheets = google.sheets("v4");

const server = new Server(
  {
    name: "example-servers/gdrive",
    version: "0.1.0",
  },
  {
    capabilities: {
      resources: {},
      tools: {},
    },
  },
);

server.setRequestHandler(ListResourcesRequestSchema, async (request) => {
  const pageSize = 10;
  const params: any = {
    pageSize,
    fields: "nextPageToken, files(id, name, mimeType)",
  };

  if (request.params?.cursor) {
    params.pageToken = request.params.cursor;
  }

  const res = await drive.files.list(params);
  const files = res.data.files!;

  return {
    resources: files.map((file) => ({
      uri: `gdrive:///${file.id}`,
      mimeType: file.mimeType,
      name: file.name,
    })),
    nextCursor: res.data.nextPageToken,
  };
});

server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
  const fileId = request.params.uri.replace("gdrive:///", "");

  const file = await drive.files.get({
    fileId,
    fields: "mimeType",
  });

  if (file.data.mimeType?.startsWith("application/vnd.google-apps")) {
    let exportMimeType: string;
    let content = '';

    if (file.data.mimeType === "application/vnd.google-apps.spreadsheet") {
      const response = await sheets.spreadsheets.get({
        spreadsheetId: fileId,
        includeGridData: true,
      });
      
      const sheetData = (response.data as any).sheets?.[0]?.data?.[0]?.rowData as RowData[] | undefined;
      if (sheetData) {
        content = sheetData.map((row: RowData) => {
          return row.values?.map(cell => cell.formattedValue || '').join(',') || '';
        }).join('\n');
      }
      exportMimeType = "text/csv";
      
      return {
        contents: [
          {
            uri: request.params.uri,
            mimeType: exportMimeType,
            text: content,
          },
        ],
      };
    }

    switch (file.data.mimeType) {
      case "application/vnd.google-apps.document":
        exportMimeType = "text/markdown";
        break;
      case "application/vnd.google-apps.presentation":
        exportMimeType = "text/plain";
        break;
      case "application/vnd.google-apps.drawing":
        exportMimeType = "image/png";
        break;
      default:
        exportMimeType = "text/plain";
    }

    const res = await drive.files.export(
      { fileId, mimeType: exportMimeType },
      { responseType: "text" },
    );

    return {
      contents: [
        {
          uri: request.params.uri,
          mimeType: exportMimeType,
          text: res.data,
        },
      ],
    };
  }

  // For regular files download content
  const res = await drive.files.get(
    { fileId, alt: "media" },
    { responseType: "arraybuffer" },
  );
  const mimeType = file.data.mimeType || "application/octet-stream";
  if (mimeType.startsWith("text/") || mimeType === "application/json") {
    return {
      contents: [
        {
          uri: request.params.uri,
          mimeType: mimeType,
          text: Buffer.from(res.data as ArrayBuffer).toString("utf-8"),
        },
      ],
    };
  } else {
    return {
      contents: [
        {
          uri: request.params.uri,
          mimeType: mimeType,
          blob: Buffer.from(res.data as ArrayBuffer).toString("base64"),
        },
      ],
    };
  }
});

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: "search",
        description: "Search for files in Google Drive",
        inputSchema: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "Search query",
            },
          },
          required: ["query"],
        },
      },
      {
        name: "list_sheets",
        description: "List all sheets in a Google Spreadsheet",
        inputSchema: {
          type: "object",
          properties: {
            fileId: {
              type: "string",
              description: "ID of the Google Spreadsheet",
            },
          },
          required: ["fileId"],
        },
      },
      {
        name: "read_sheet",
        description: "Read data from a specific sheet in a Google Spreadsheet",
        inputSchema: {
          type: "object",
          properties: {
            fileId: {
              type: "string",
              description: "ID of the Google Spreadsheet",
            },
            sheetName: {
              type: "string",
              description: "Name of the sheet to read (optional, defaults to first sheet)",
            },
            range: {
              type: "string",
              description: "Range to read in A1 notation (optional, defaults to entire sheet)",
            },
          },
          required: ["fileId"],
        },
      },
      {
        name: "edit_sheet",
        description: "Edit data in a specific sheet in a Google Spreadsheet",
        inputSchema: {
          type: "object",
          properties: {
            fileId: {
              type: "string",
              description: "ID of the Google Spreadsheet",
            },
            sheetName: {
              type: "string",
              description: "Name of the sheet to edit (optional, defaults to first sheet)",
            },
            range: {
              type: "string",
              description: "Range to update in A1 notation (required)",
            },
            values: {
              type: "array",
              description: "2D array of values to update in the specified range",
              items: {
                type: "array",
                items: {}
              }
            },
          },
          required: ["fileId", "range", "values"],
        },
      },
      {
        name: "format_sheet",
        description: "Apply formatting (colors, fonts, etc.) to cells in a Google Spreadsheet",
        inputSchema: {
          type: "object",
          properties: {
            fileId: {
              type: "string",
              description: "ID of the Google Spreadsheet",
            },
            sheetName: {
              type: "string",
              description: "Name of the sheet to format (optional, defaults to first sheet)",
            },
            range: {
              type: "string",
              description: "Range to format in A1 notation (required)",
            },
            formatting: {
              type: "object",
              description: "Formatting options to apply",
              properties: {
                backgroundColor: {
                  type: "string",
                  description: "Background color in hex format (e.g. #ff0000 for red)"
                },
                textColor: {
                  type: "string",
                  description: "Text color in hex format (e.g. #0000ff for blue)"
                },
                fontSize: {
                  type: "number",
                  description: "Font size in points"
                },
                bold: {
                  type: "boolean",
                  description: "Whether text should be bold"
                },
                italic: {
                  type: "boolean",
                  description: "Whether text should be italic"
                },
                horizontalAlignment: {
                  type: "string",
                  description: "Horizontal alignment (LEFT, CENTER, RIGHT)"
                },
                verticalAlignment: {
                  type: "string",
                  description: "Vertical alignment (TOP, MIDDLE, BOTTOM)"
                }
              }
            }
          },
          required: ["fileId", "range", "formatting"],
        },
      },
    ],
  };
});

// Thêm hàm helper để xử lý URL và ID
function extractFileIdFromUrl(input: string): string {
  if (!input.includes('http')) return input;

  try {
    const url = new URL(input);
    // Xử lý các dạng URL khác nhau
    // 1. /spreadsheets/d/[ID]/edit
    const d = url.pathname.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (d) return d[1];

    // 2. /spreadsheets/d/[ID]
    const spreadsheet = url.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (spreadsheet) return spreadsheet[1];

    // 3. /open?id=[ID]
    const id = url.searchParams.get('id');
    if (id) return id;

    throw new Error('Không thể trích xuất ID từ URL');
  } catch (err) {
    throw new Error('URL không hợp lệ');
  }
}

// Cập nhật hàm verifyFileAccess để kiểm tra quyền chỉnh sửa
async function verifyFileAccess(fileId: string, requireEdit: boolean = false): Promise<boolean> {
  try {
    const file = await drive.files.get({
      fileId,
      fields: "id,mimeType,capabilities(canReadRevisions,canEdit),permissions",
    });

    if (!file.data.mimeType?.includes('spreadsheet')) {
      throw new Error('File không phải là Google Spreadsheet');
    }

    // Kiểm tra chi tiết về quyền truy cập
    const permissions = await drive.permissions.list({
      fileId: fileId,
      fields: 'permissions(role,type,emailAddress)'
    });

    // Kiểm tra xem file có được chia sẻ công khai không
    const isPublic = permissions.data.permissions?.some(
      p => p.type === 'anyone' || p.type === 'domain'
    );

    // Kiểm tra xem người dùng hiện tại có quyền truy cập không
    const hasReadAccess = file.data.capabilities?.canReadRevisions === true;
    const hasEditAccess = file.data.capabilities?.canEdit === true;

    if (requireEdit && !hasEditAccess) {
      const userEmail = await getCurrentUserEmail();
      throw new Error(
        'Không có quyền chỉnh sửa file.\n' +
        `Email đang sử dụng: ${userEmail}\n` +
        'Vui lòng:\n' +
        '1. Đảm bảo file được chia sẻ với email của bạn với quyền chỉnh sửa\n' +
        '2. Hoặc đặt chế độ chia sẻ "Anyone with the link can edit"\n' +
        '3. Thử xác thực lại bằng lệnh: node dist/index.js auth'
      );
    }

    if (!isPublic && !hasReadAccess) {
      const userEmail = await getCurrentUserEmail();
      throw new Error(
        'Không có quyền truy cập file.\n' +
        `Email đang sử dụng: ${userEmail}\n` +
        'Vui lòng:\n' +
        '1. Đảm bảo file được chia sẻ với email của bạn\n' +
        '2. Hoặc đặt chế độ chia sẻ "Anyone with the link can view"\n' +
        '3. Thử xác thực lại bằng lệnh: node dist/index.js auth'
      );
    }

    return true;
  } catch (err: any) {
    if (err.code === 404) {
      const userEmail = await getCurrentUserEmail();
      throw new Error(
        'File không tồn tại hoặc bạn không có quyền truy cập.\n' +
        `Email đang sử dụng: ${userEmail}\n` +
        'Vui lòng kiểm tra:\n' +
        '1. URL hoặc ID file chính xác\n' +
        '2. File đã được chia sẻ với email của bạn\n' +
        '3. Bạn đã xác thực với đúng tài khoản Google'
      );
    }
    throw err;
  }
}

// Thêm hàm để lấy email người dùng hiện tại
async function getCurrentUserEmail(): Promise<string> {
  try {
    const about = await drive.about.get({
      fields: 'user(emailAddress)'
    });
    return about.data.user?.emailAddress || '';
  } catch (err) {
    console.error('Không thể lấy thông tin người dùng:', err);
    return '';
  }
}

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  if (request.params.name === "search") {
    const userQuery = request.params.arguments?.query as string;
    const escapedQuery = userQuery.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    const formattedQuery = `fullText contains '${escapedQuery}'`;

    const res = await drive.files.list({
      q: formattedQuery,
      pageSize: 10,
      fields: "files(id, name, mimeType, modifiedTime, size)",
    });

    const fileList = res.data.files
      ?.map((file: any) => `${file.name} (${file.mimeType})`)
      .join("\n");
    return {
      content: [
        {
          type: "text",
          text: `Found ${res.data.files?.length ?? 0} files:\n${fileList}`,
        },
      ],
      isError: false,
    };
  }

  if (request.params.name === "list_sheets") {
    try {
      const inputId = request.params.arguments?.fileId as string;
      const fileId = extractFileIdFromUrl(inputId);
      
      await verifyFileAccess(fileId);

      const response = await sheets.spreadsheets.get({
        spreadsheetId: fileId,
        fields: "sheets.properties",
      }) as SpreadsheetResponse;

      const sheetsList = response.data.sheets?.map((sheet: Sheet) => ({
        name: sheet.properties?.title,
        id: sheet.properties?.sheetId,
        rowCount: sheet.properties?.gridProperties?.rowCount,
        columnCount: sheet.properties?.gridProperties?.columnCount,
      }));

      return {
        content: [
          {
            type: "text",
            text: `Danh sách các sheet:\n${JSON.stringify(sheetsList, null, 2)}`,
          },
        ],
        isError: false,
      };
    } catch (err) {
      const error = err as Error;
      return {
        content: [
          {
            type: "text",
            text: `Không thể lấy danh sách sheet. Lỗi: ${error.message}\n` +
                  `Vui lòng kiểm tra:\n` +
                  `1. URL hoặc ID chính xác\n` +
                  `2. File là Google Spreadsheet\n` +
                  `3. File đã được chia sẻ với bạn`,
          },
        ],
        isError: true,
      };
    }
  }

  if (request.params.name === "read_sheet") {
    try {
      const inputId = request.params.arguments?.fileId as string;
      const fileId = extractFileIdFromUrl(inputId);
      
      // Lấy email người dùng để hiển thị trong thông báo lỗi
      const userEmail = await getCurrentUserEmail();
      
      try {
        await verifyFileAccess(fileId);
      } catch (err) {
        const error = err as Error;
        return {
          content: [
            {
              type: "text",
              text: `${error.message}\n\n` +
                    `Email đang sử dụng: ${userEmail}\n` +
                    `File ID: ${fileId}`
            },
          ],
          isError: true,
        };
      }

      const sheetName = request.params.arguments?.sheetName as string | undefined;
      const rangeInput = request.params.arguments?.range as string | undefined;

      let actualRange: string | undefined = rangeInput;
      if (sheetName && !rangeInput) {
        const metadata = await sheets.spreadsheets.get({
          spreadsheetId: fileId,
          fields: "sheets.properties",
        }) as SpreadsheetResponse;
        
        const sheet = metadata.data.sheets?.find((s: Sheet) => 
          s.properties?.title === sheetName
        );
        
        if (!sheet) {
          throw new Error(`Không tìm thấy sheet có tên "${sheetName}"`);
        }
        
        actualRange = `${sheetName}!A1:${String.fromCharCode(64 + (sheet.properties?.gridProperties?.columnCount || 26))}${sheet.properties?.gridProperties?.rowCount || 1000}`;
      }

      const effectiveRange: string = actualRange || sheetName || 'Sheet1!A1:Z1000';

      const requestParams: sheets_v4.Params$Resource$Spreadsheets$Values$Get = {
        spreadsheetId: fileId,
        range: effectiveRange,
        valueRenderOption: 'FORMATTED_VALUE',
        majorDimension: 'ROWS'
      };

      const response = await sheets.spreadsheets.values.get(requestParams);
      
      if (!response.data || !response.data.values || response.data.values.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "Sheet không có dữ liệu",
            },
          ],
          isError: false,
        };
      }

      const formattedData = response.data.values.map((row: any[]) => 
        row.map((cell: any) => (cell?.toString() || '')).join(',')
      ).join('\n');
      
      return {
        content: [
          {
            type: "text",
            text: `Dữ liệu sheet:\n${formattedData}`,
          },
        ],
        isError: false,
      };
    } catch (err) {
      const error = err as Error;
      return {
        content: [
          {
            type: "text", 
            text: `Không thể đọc sheet. Lỗi: ${error.message}\n` +
                  `Vui lòng kiểm tra:\n` +
                  `1. URL hoặc ID chính xác\n` +
                  `2. File là Google Spreadsheet\n` +
                  `3. File đã được chia sẻ với bạn\n` +
                  `4. Tên sheet chính xác (nếu có cung cấp)`,
          },
        ],
        isError: true,
      };
    }
  }

  if (request.params.name === "edit_sheet") {
    try {
      const inputId = request.params.arguments?.fileId as string;
      const fileId = extractFileIdFromUrl(inputId);
      
      // Lấy email người dùng để hiển thị trong thông báo lỗi
      const userEmail = await getCurrentUserEmail();
      
      try {
        // Kiểm tra xem có quyền edit không
        await verifyFileAccess(fileId, true);
      } catch (err) {
        const error = err as Error;
        return {
          content: [
            {
              type: "text",
              text: `${error.message}\n\n` +
                    `Email đang sử dụng: ${userEmail}\n` +
                    `File ID: ${fileId}`
            },
          ],
          isError: true,
        };
      }

      const sheetName = request.params.arguments?.sheetName as string | undefined;
      const rangeInput = request.params.arguments?.range as string;
      const values = request.params.arguments?.values as any[][];

      if (!values || !Array.isArray(values)) {
        throw new Error('Dữ liệu cung cấp không hợp lệ. Cần phải là mảng 2 chiều.');
      }

      // Xác định range thực tế nếu sheetName được cung cấp
      let actualRange = rangeInput;
      if (sheetName && !rangeInput.includes('!')) {
        actualRange = `${sheetName}!${rangeInput}`;
      }

      // Nếu range không có sheetName, thêm 'Sheet1!' vào trước nếu cần
      if (!actualRange.includes('!')) {
        actualRange = `Sheet1!${actualRange}`;
      }

      // Tiến hành cập nhật dữ liệu
      const updateResponse = await sheets.spreadsheets.values.update({
        spreadsheetId: fileId,
        range: actualRange,
        valueInputOption: 'USER_ENTERED', // Xử lý công thức và định dạng số
        requestBody: {
          values: values
        }
      });

      if (updateResponse.status !== 200) {
        throw new Error(`Cập nhật thất bại. Mã trạng thái: ${updateResponse.status}`);
      }

      return {
        content: [
          {
            type: "text",
            text: `Cập nhật thành công!\n` +
                  `- Sheet: ${actualRange.split('!')[0]}\n` +
                  `- Range: ${actualRange.split('!')[1]}\n` +
                  `- Số hàng cập nhật: ${updateResponse.data.updatedRows}\n` +
                  `- Số cột cập nhật: ${updateResponse.data.updatedColumns}\n` +
                  `- Số ô cập nhật: ${updateResponse.data.updatedCells}`
          },
        ],
        isError: false,
      };
    } catch (err) {
      const error = err as Error;
      return {
        content: [
          {
            type: "text", 
            text: `Không thể cập nhật sheet. Lỗi: ${error.message}\n` +
                  `Vui lòng kiểm tra:\n` +
                  `1. URL hoặc ID chính xác\n` +
                  `2. File là Google Spreadsheet\n` +
                  `3. File đã được chia sẻ với bạn với quyền chỉnh sửa\n` +
                  `4. Tên sheet và range chính xác\n` +
                  `5. Dữ liệu cập nhật theo đúng định dạng mảng 2 chiều`,
          },
        ],
        isError: true,
      };
    }
  }

  if (request.params.name === "format_sheet") {
    try {
      const inputId = request.params.arguments?.fileId as string;
      const fileId = extractFileIdFromUrl(inputId);
      
      // Lấy email người dùng để hiển thị trong thông báo lỗi
      const userEmail = await getCurrentUserEmail();
      
      try {
        // Kiểm tra xem có quyền edit không
        await verifyFileAccess(fileId, true);
      } catch (err) {
        const error = err as Error;
        return {
          content: [
            {
              type: "text",
              text: `${error.message}\n\n` +
                    `Email đang sử dụng: ${userEmail}\n` +
                    `File ID: ${fileId}`
            },
          ],
          isError: true,
        };
      }

      const sheetName = request.params.arguments?.sheetName as string | undefined;
      const rangeInput = request.params.arguments?.range as string;
      const formatting = request.params.arguments?.formatting as any;

      if (!formatting || typeof formatting !== 'object') {
        throw new Error('Định dạng cung cấp không hợp lệ.');
      }

      // Xác định range thực tế nếu sheetName được cung cấp
      let actualRange = rangeInput;
      if (sheetName && !rangeInput.includes('!')) {
        actualRange = `${sheetName}!${rangeInput}`;
      }

      // Nếu range không có sheetName, thêm 'Sheet1!' vào trước nếu cần
      if (!actualRange.includes('!')) {
        actualRange = `Sheet1!${actualRange}`;
      }

      // Lấy thông tin về sheet và grid location
      const spreadsheet = await sheets.spreadsheets.get({
        spreadsheetId: fileId,
        ranges: [actualRange],
        fields: 'sheets.properties,sheets.merges'
      });

      const sheetId = spreadsheet.data.sheets?.[0].properties?.sheetId;
      if (!sheetId) {
        throw new Error('Không thể xác định sheet ID');
      }

      // Phân tích range để lấy start row/col và end row/col
      const rangeParts = actualRange.split('!')[1].match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      let startCol, startRow, endCol, endRow;
      
      if (rangeParts) {
        startCol = columnLetterToNumber(rangeParts[1]);
        startRow = parseInt(rangeParts[2]) - 1; // 0-based index
        endCol = columnLetterToNumber(rangeParts[3]);
        endRow = parseInt(rangeParts[4]) - 1; // 0-based index
      } else {
        // Single cell
        const cellParts = actualRange.split('!')[1].match(/([A-Z]+)(\d+)/);
        if (!cellParts) {
          throw new Error('Định dạng range không hợp lệ');
        }
        startCol = columnLetterToNumber(cellParts[1]);
        startRow = parseInt(cellParts[2]) - 1;
        endCol = startCol;
        endRow = startRow;
      }

      // Tạo request để cập nhật định dạng
      const requests = [];
      
      // Tạo CellFormat từ formatting options
      const userFormatting: any = {};
      
      if (formatting.backgroundColor) {
        const color = hexToRgb(formatting.backgroundColor);
        userFormatting.backgroundColor = {
          red: color.r / 255,
          green: color.g / 255,
          blue: color.b / 255,
          alpha: 1
        };
      }
      
      if (formatting.textColor) {
        const color = hexToRgb(formatting.textColor);
        userFormatting.textFormat = {
          ...userFormatting.textFormat,
          foregroundColor: {
            red: color.r / 255,
            green: color.g / 255,
            blue: color.b / 255,
            alpha: 1
          }
        };
      }
      
      if (formatting.fontSize) {
        userFormatting.textFormat = {
          ...userFormatting.textFormat,
          fontSize: formatting.fontSize
        };
      }
      
      if (formatting.bold !== undefined) {
        userFormatting.textFormat = {
          ...userFormatting.textFormat,
          bold: formatting.bold
        };
      }
      
      if (formatting.italic !== undefined) {
        userFormatting.textFormat = {
          ...userFormatting.textFormat,
          italic: formatting.italic
        };
      }
      
      if (formatting.horizontalAlignment) {
        userFormatting.horizontalAlignment = formatting.horizontalAlignment.toUpperCase();
      }
      
      if (formatting.verticalAlignment) {
        userFormatting.verticalAlignment = formatting.verticalAlignment.toUpperCase();
      }
      
      // Thêm request cập nhật định dạng
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: startRow,
            endRowIndex: endRow + 1,
            startColumnIndex: startCol,
            endColumnIndex: endCol + 1
          },
          cell: {
            userEnteredFormat: userFormatting
          },
          fields: 'userEnteredFormat(' + Object.keys(userFormatting).join(',') + ')'
        }
      });

      // Thực hiện batch update
      const updateResponse = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: fileId,
        requestBody: {
          requests: requests
        }
      });

      if (updateResponse.status !== 200) {
        throw new Error(`Cập nhật định dạng thất bại. Mã trạng thái: ${updateResponse.status}`);
      }

      return {
        content: [
          {
            type: "text",
            text: `Cập nhật định dạng thành công!\n` +
                  `- Sheet: ${actualRange.split('!')[0]}\n` +
                  `- Range: ${actualRange.split('!')[1]}\n` +
                  `- Các thuộc tính đã áp dụng: ${Object.keys(formatting).join(', ')}`
          },
        ],
        isError: false,
      };
    } catch (err) {
      const error = err as Error;
      return {
        content: [
          {
            type: "text", 
            text: `Không thể cập nhật định dạng. Lỗi: ${error.message}\n` +
                  `Vui lòng kiểm tra:\n` +
                  `1. URL hoặc ID chính xác\n` +
                  `2. File là Google Spreadsheet\n` +
                  `3. File đã được chia sẻ với bạn với quyền chỉnh sửa\n` +
                  `4. Tên sheet và range chính xác\n` +
                  `5. Định dạng cung cấp hợp lệ`,
          },
        ],
        isError: true,
      };
    }
  }

  throw new Error("Tool not found");
});

// Hàm helper chuyển cột dạng chữ thành số (A->0, B->1, etc.)
function columnLetterToNumber(column: string): number {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result - 1; // 0-based index
}

// Hàm chuyển đổi màu hex sang RGB
function hexToRgb(hex: string): { r: number, g: number, b: number } {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : { r: 0, g: 0, b: 0 };
}

const credentialsPath = process.env.GDRIVE_CREDENTIALS_PATH || path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../.gdrive-server-credentials.json",
);

async function authenticateAndSaveCredentials() {
  console.log("Launching auth flow…");
  const auth = await authenticate({
    keyfilePath: process.env.GDRIVE_OAUTH_PATH || path.join(
      path.dirname(fileURLToPath(import.meta.url)),
      "../../../gcp-oauth.keys.json",
    ),
    scopes: [
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/spreadsheets.readonly",
      "https://www.googleapis.com/auth/drive.file",
      "https://www.googleapis.com/auth/spreadsheets",
    ],
  });
  fs.writeFileSync(credentialsPath, JSON.stringify(auth.credentials));
  console.log("Credentials saved. You can now run the server.");
}

async function loadCredentialsAndRunServer() {
  if (!fs.existsSync(credentialsPath)) {
    console.error(
      "Credentials not found. Please run with 'auth' argument first.",
    );
    process.exit(1);
  }

  const credentials = JSON.parse(fs.readFileSync(credentialsPath, "utf-8"));
  const auth = new google.auth.OAuth2();
  auth.setCredentials(credentials);
  google.options({ auth });

  console.error("Credentials loaded. Starting server.");
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

if (process.argv[2] === "auth") {
  authenticateAndSaveCredentials().catch(console.error);
} else {
  loadCredentialsAndRunServer().catch(console.error);
}
