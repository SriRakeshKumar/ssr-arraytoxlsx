"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.HeaderColorThemes = void 0;
exports.ArrayToXLSX = ArrayToXLSX;
exports.csvToArray = csvToArray;
exports.validateArrayData = validateArrayData;
exports.validateHexColor = validateHexColor;
exports.createHeaderStyle = createHeaderStyle;
const XLSX = __importStar(require("xlsx-js-style"));
/**
 * Predefined color themes for headers
 */
exports.HeaderColorThemes = {
    blue: { backgroundColor: '#4472C4', textColor: '#FFFFFF', bold: true },
    green: { backgroundColor: '#70AD47', textColor: '#FFFFFF', bold: true },
    red: { backgroundColor: '#E15759', textColor: '#FFFFFF', bold: true },
    orange: { backgroundColor: '#F79646', textColor: '#FFFFFF', bold: true },
    purple: { backgroundColor: '#9F4F96', textColor: '#FFFFFF', bold: true },
    teal: { backgroundColor: '#4BACC6', textColor: '#FFFFFF', bold: true },
    gray: { backgroundColor: '#A5A5A5', textColor: '#FFFFFF', bold: true },
    darkBlue: { backgroundColor: '#2F5597', textColor: '#FFFFFF', bold: true },
    darkGreen: { backgroundColor: '#548235', textColor: '#FFFFFF', bold: true },
    corporate: { backgroundColor: '#1F4E79', textColor: '#FFFFFF', bold: true },
};
/**
 * Downloads array data as an XLSX file with customizable header styling
 * @param data - Array of objects or array of arrays to be converted
 * @param options - Configuration options for the download
 */
function ArrayToXLSX(data, options = {}) {
    if (!data || !Array.isArray(data) || data.length === 0) {
        throw new Error('Data must be a non-empty array');
    }
    const { filename = 'report', sheetName = 'Sheet1', headers, includeHeaders = true, headerStyle, autoFitColumns = true } = options;
    try {
        // Create a new workbook
        const workbook = XLSX.utils.book_new();
        let worksheet;
        let headerRange = null;
        // Handle array of objects
        if (data.length > 0 && typeof data[0] === 'object' && !Array.isArray(data[0])) {
            const objectData = data;
            const keys = Object.keys(objectData[0]);
            if (headers && includeHeaders) {
                // Use custom headers
                const mappedData = objectData.map(row => {
                    const mappedRow = {};
                    headers.forEach((header, index) => {
                        mappedRow[header] = row[keys[index]] || '';
                    });
                    return mappedRow;
                });
                worksheet = XLSX.utils.json_to_sheet(mappedData);
                headerRange = `A1:${XLSX.utils.encode_col(headers.length - 1)}1`;
            }
            else if (!includeHeaders) {
                // Convert to array of arrays without headers
                const arrayData = objectData.map(row => keys.map(key => row[key]));
                worksheet = XLSX.utils.aoa_to_sheet(arrayData);
            }
            else {
                // Use default object keys as headers
                worksheet = XLSX.utils.json_to_sheet(objectData);
                headerRange = `A1:${XLSX.utils.encode_col(keys.length - 1)}1`;
            }
        }
        // Handle array of arrays
        else if (Array.isArray(data[0])) {
            const arrayData = data;
            if (headers && includeHeaders) {
                // Add headers as first row
                const dataWithHeaders = [headers, ...arrayData];
                worksheet = XLSX.utils.aoa_to_sheet(dataWithHeaders);
                headerRange = `A1:${XLSX.utils.encode_col(headers.length - 1)}1`;
            }
            else if (includeHeaders && arrayData.length > 0) {
                // Treat first row as headers
                worksheet = XLSX.utils.aoa_to_sheet(arrayData);
                headerRange = `A1:${XLSX.utils.encode_col(arrayData[0].length - 1)}1`;
            }
            else {
                worksheet = XLSX.utils.aoa_to_sheet(arrayData);
            }
        }
        else {
            throw new Error('Unsupported data format. Data should be an array of objects or array of arrays.');
        }
        // Apply header styling if provided
        if (headerStyle && headerRange && includeHeaders) {
            applyHeaderStyling(worksheet, headerRange, headerStyle);
        }
        // Auto-fit columns if requested
        if (autoFitColumns) {
            autoFitWorksheetColumns(worksheet);
        }
        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        // Generate XLSX file buffer
        const xlsxBuffer = XLSX.write(workbook, {
            bookType: 'xlsx',
            type: 'array',
            compression: true,
            cellStyles: true // Enable cell styling
        });
        // Create blob and download
        const blob = new Blob([xlsxBuffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        downloadBlob(blob, `${filename}.xlsx`);
    }
    catch (error) {
        throw new Error(`Failed to create XLSX file: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
}
/**
 * Apply styling to header cells
 * @param worksheet - The worksheet to apply styling to
 * @param headerRange - The range of header cells (e.g., 'A1:D1')
 * @param headerStyle - The styling options to apply
 */
function applyHeaderStyling(worksheet, headerRange, headerStyle) {
    const range = XLSX.utils.decode_range(headerRange);
    for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
        if (!worksheet[cellAddress])
            continue;
        // Initialize cell style if it doesn't exist
        if (!worksheet[cellAddress].s) {
            worksheet[cellAddress].s = {};
        }
        const cellStyle = worksheet[cellAddress].s;
        // Apply background color
        if (headerStyle.backgroundColor) {
            const bgColor = headerStyle.backgroundColor.replace('#', '');
            cellStyle.fill = {
                fgColor: { rgb: bgColor.toUpperCase() }
            };
        }
        // Apply text color and font styling
        if (headerStyle.textColor || headerStyle.bold || headerStyle.fontSize) {
            cellStyle.font = {
                ...(cellStyle.font || {}),
                ...(headerStyle.textColor && {
                    color: { rgb: headerStyle.textColor.replace('#', '').toUpperCase() }
                }),
                ...(headerStyle.bold && { bold: true }),
                ...(headerStyle.fontSize && { sz: headerStyle.fontSize })
            };
        }
        // Apply borders for a more professional look
        cellStyle.border = {
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
        };
        // Center align headers
        cellStyle.alignment = {
            horizontal: 'center',
            vertical: 'center'
        };
    }
}
/**
 * Auto-fit column widths based on content
 * @param worksheet - The worksheet to auto-fit
 */
function autoFitWorksheetColumns(worksheet) {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    const colWidths = [];
    // Calculate maximum width for each column
    for (let col = range.s.c; col <= range.e.c; col++) {
        let maxWidth = 10; // Minimum width
        for (let row = range.s.r; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            if (cell && cell.v) {
                const cellValue = String(cell.v);
                const cellWidth = cellValue.length + 2; // Add some padding
                maxWidth = Math.max(maxWidth, Math.min(cellWidth, 50)); // Cap at 50 characters
            }
        }
        colWidths[col] = maxWidth;
    }
    // Apply column widths
    worksheet['!cols'] = colWidths.map(width => ({ width }));
}
/**
 * Helper function to download a blob as a file
 * @param blob - The blob to download
 * @param filename - The filename for the download
 */
function downloadBlob(blob, filename) {
    if (typeof window === 'undefined') {
        throw new Error('Download functionality is only available in browser environments');
    }
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    // Append to body, click, and remove
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    // Release the object URL
    window.URL.revokeObjectURL(url);
}
/**
 * Utility function to convert CSV string to array format
 * @param csvString - CSV formatted string
 * @param delimiter - CSV delimiter (default: ',')
 * @returns Array of arrays
 */
function csvToArray(csvString, delimiter = ',') {
    const lines = csvString.trim().split('\n');
    return lines.map(line => line.split(delimiter).map(cell => cell.trim()));
}
/**
 * Utility function to validate data before conversion
 * @param data - Data to validate
 * @returns boolean indicating if data is valid
 */
function validateArrayData(data) {
    return Array.isArray(data) &&
        data.length > 0 &&
        (Array.isArray(data[0]) || (typeof data[0] === 'object' && data[0] !== null));
}
/**
 * Utility function to validate hex color format
 * @param color - Color string to validate
 * @returns boolean indicating if color is valid
 */
function validateHexColor(color) {
    const hexColorRegex = /^#?([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/;
    return hexColorRegex.test(color);
}
/**
 * Helper function to create custom header style
 * @param backgroundColor - Background color in hex
 * @param textColor - Text color in hex (optional, defaults to white)
 * @param bold - Whether text should be bold (optional, defaults to true)
 * @returns HeaderColorOptions object
 */
function createHeaderStyle(backgroundColor, textColor = '#FFFFFF', bold = true, fontSize = 12) {
    if (!validateHexColor(backgroundColor)) {
        throw new Error('Invalid background color format. Use hex format like #FF0000 or FF0000');
    }
    if (!validateHexColor(textColor)) {
        throw new Error('Invalid text color format. Use hex format like #FFFFFF or FFFFFF');
    }
    return {
        backgroundColor,
        textColor,
        bold,
        fontSize
    };
}
// Default export for convenience
exports.default = ArrayToXLSX;
//# sourceMappingURL=index.js.map