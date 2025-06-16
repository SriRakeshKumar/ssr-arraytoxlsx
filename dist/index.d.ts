/**
 * Configuration options for downloading XLSX files
 */
export interface DownloadOptions {
    /** The filename for the downloaded file (without extension) */
    filename?: string;
    /** The worksheet name */
    sheetName?: string;
    /** Column headers (if not provided, will use object keys or array indices) */
    headers?: string[];
    /** Whether to include headers in the output */
    includeHeaders?: boolean;
}
/**
 * Type for array data that can be converted to XLSX
 */
export type ArrayData = Array<Record<string, any>> | Array<Array<any>>;
/**
 * Downloads array data as an XLSX file
 * @param data - Array of objects or array of arrays to be converted
 * @param options - Configuration options for the download
 */
export declare function downloadArrayAsXLSX(data: ArrayData, options?: DownloadOptions): void;
/**
 * Utility function to convert CSV string to array format
 * @param csvString - CSV formatted string
 * @param delimiter - CSV delimiter (default: ',')
 * @returns Array of arrays
 */
export declare function csvToArray(csvString: string, delimiter?: string): string[][];
/**
 * Utility function to validate data before conversion
 * @param data - Data to validate
 * @returns boolean indicating if data is valid
 */
export declare function validateArrayData(data: any): data is ArrayData;
export default downloadArrayAsXLSX;
//# sourceMappingURL=index.d.ts.map