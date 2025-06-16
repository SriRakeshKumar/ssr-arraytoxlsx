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
exports.downloadArrayAsXLSX = downloadArrayAsXLSX;
exports.csvToArray = csvToArray;
exports.validateArrayData = validateArrayData;
const XLSX = __importStar(require("xlsx"));
/**
 * Downloads array data as an XLSX file
 * @param data - Array of objects or array of arrays to be converted
 * @param options - Configuration options for the download
 */
function downloadArrayAsXLSX(data, options = {}) {
    if (!data || !Array.isArray(data) || data.length === 0) {
        throw new Error('Data must be a non-empty array');
    }
    const { filename = 'report', sheetName = 'Sheet1', headers, includeHeaders = true } = options;
    try {
        // Create a new workbook
        const workbook = XLSX.utils.book_new();
        let worksheet;
        // Handle array of objects
        if (data.length > 0 && typeof data[0] === 'object' && !Array.isArray(data[0])) {
            const objectData = data;
            if (headers && includeHeaders) {
                // Use custom headers
                const mappedData = objectData.map(row => {
                    const mappedRow = {};
                    headers.forEach((header, index) => {
                        const keys = Object.keys(row);
                        mappedRow[header] = row[keys[index]] || '';
                    });
                    return mappedRow;
                });
                worksheet = XLSX.utils.json_to_sheet(mappedData);
            }
            else if (!includeHeaders) {
                // Convert to array of arrays without headers
                const keys = Object.keys(objectData[0]);
                const arrayData = objectData.map(row => keys.map(key => row[key]));
                worksheet = XLSX.utils.aoa_to_sheet(arrayData);
            }
            else {
                // Use default object keys as headers
                worksheet = XLSX.utils.json_to_sheet(objectData);
            }
        }
        // Handle array of arrays
        else if (Array.isArray(data[0])) {
            const arrayData = data;
            if (headers && includeHeaders) {
                // Add headers as first row
                const dataWithHeaders = [headers, ...arrayData];
                worksheet = XLSX.utils.aoa_to_sheet(dataWithHeaders);
            }
            else {
                worksheet = XLSX.utils.aoa_to_sheet(arrayData);
            }
        }
        else {
            throw new Error('Unsupported data format. Data should be an array of objects or array of arrays.');
        }
        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        // Generate XLSX file buffer
        const xlsxBuffer = XLSX.write(workbook, {
            bookType: 'xlsx',
            type: 'array',
            compression: true
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
// Default export for convenience
exports.default = downloadArrayAsXLSX;
//# sourceMappingURL=index.js.map