/**
 * Color configuration for headers
 */
export interface HeaderColorOptions {
    /** Background color in hex format (e.g., '#FF0000' or 'FF0000') */
    backgroundColor?: string;
    /** Text color in hex format (e.g., '#FFFFFF' or 'FFFFFF') */
    textColor?: string;
    /** Whether text should be bold */
    bold?: boolean;
    /** Font size for headers */
    fontSize?: number;
}
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
    /** Color and styling options for headers */
    headerStyle?: HeaderColorOptions;
    /** Auto-fit column widths based on content */
    autoFitColumns?: boolean;
}
/**
 * Type for array data that can be converted to XLSX
 */
export type ArrayData = Array<Record<string, any>> | Array<Array<any>>;
/**
 * Predefined color themes for headers
 */
export declare const HeaderColorThemes: {
    readonly blue: {
        readonly backgroundColor: "#4472C4";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
    readonly green: {
        readonly backgroundColor: "#70AD47";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
    readonly red: {
        readonly backgroundColor: "#E15759";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
    readonly orange: {
        readonly backgroundColor: "#F79646";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
    readonly purple: {
        readonly backgroundColor: "#9F4F96";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
    readonly teal: {
        readonly backgroundColor: "#4BACC6";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
    readonly gray: {
        readonly backgroundColor: "#A5A5A5";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
    readonly darkBlue: {
        readonly backgroundColor: "#2F5597";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
    readonly darkGreen: {
        readonly backgroundColor: "#548235";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
    readonly corporate: {
        readonly backgroundColor: "#1F4E79";
        readonly textColor: "#FFFFFF";
        readonly bold: true;
    };
};
/**
 * Downloads array data as an XLSX file with customizable header styling
 * @param data - Array of objects or array of arrays to be converted
 * @param options - Configuration options for the download
 */
export declare function ArrayToXLSX(data: ArrayData, options?: DownloadOptions): void;
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
/**
 * Utility function to validate hex color format
 * @param color - Color string to validate
 * @returns boolean indicating if color is valid
 */
export declare function validateHexColor(color: string): boolean;
/**
 * Helper function to create custom header style
 * @param backgroundColor - Background color in hex
 * @param textColor - Text color in hex (optional, defaults to white)
 * @param bold - Whether text should be bold (optional, defaults to true)
 * @returns HeaderColorOptions object
 */
export declare function createHeaderStyle(backgroundColor: string, textColor?: string, bold?: boolean, fontSize?: number): HeaderColorOptions;
export default ArrayToXLSX;
//# sourceMappingURL=index.d.ts.map