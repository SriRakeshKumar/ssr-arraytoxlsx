# Array to XLSX Downloader

A lightweight, TypeScript-compatible npm library for converting array data to downloadable XLSX (Excel) files in the browser.

> **🌐 Browser-Only Library**: This library is designed specifically for frontend applications and requires a browser environment to function. It will not work in Node.js, server-side rendering, or backend environments.

## Features

- ✅ Convert arrays to XLSX files in the browser
- ✅ TypeScript support with full type definitions
- ✅ JavaScript compatible
- ✅ Supports array of objects and array of arrays
- ✅ Customizable filename and sheet name
- ✅ Custom headers support
- ✅ Lightweight with minimal dependencies
- ✅ Client-side file download (no server required)
- ✅ Error handling and validation

## Compatibility

### ✅ Works With (Frontend/Client-Side)

- React, Vue, Angular, Svelte applications
- Vanilla JavaScript/HTML projects
- Client-side bundlers (Webpack, Vite, Parcel)
- Browser environments with modern JavaScript support

### ❌ Does Not Work With (Backend/Server-Side)

- Node.js applications
- Server-side rendering (SSR)
- Next.js server-side functions
- Express.js or other backend frameworks
- Electron main process

## Installation

```bash
npm install ssr-xlsx-export
```

## Usage

### Basic Usage (JavaScript)

```javascript
import { ArrayToXLSX } from "ssr-xlsx-export";

// Example: Download button click handler
document.getElementById("download-report").addEventListener("click", () => {
  const reportData = [
    { name: "John Doe", age: 30, city: "New York" },
    { name: "Jane Smith", age: 25, city: "Los Angeles" },
    { name: "Bob Johnson", age: 35, city: "Chicago" },
  ];

  ArrayToXLSX(reportData, {
    filename: "user-report",
    sheetName: "Users",
  });
});
```

### TypeScript Usage

```typescript
import { ArrayToXLSX, ArrayData, DownloadOptions } from "ssr-xlsx-export";

interface User {
  name: string;
  age: number;
  city: string;
}

const handleDownload = (data: User[]) => {
  const options: DownloadOptions = {
    filename: "user-report",
    sheetName: "Users",
    headers: ["Full Name", "Age", "City"],
    includeHeaders: true,
  };

  ArrayToXLSX(data, options);
};
```

### React Example

```jsx
import React, { useState } from "react";
import { ArrayToXLSX } from "ssr-xlsx-export";

const ReportComponent = () => {
  const [reportData, setReportData] = useState([]);

  const handleDownloadReport = () => {
    // Your data fetching logic here
    const data = [
      { product: "Laptop", sales: 150, revenue: 75000 },
      { product: "Phone", sales: 300, revenue: 60000 },
      { product: "Tablet", sales: 200, revenue: 40000 },
    ];

    ArrayToXLSX(data, {
      filename: "sales-report",
      sheetName: "Q1 Sales",
      headers: ["Product Name", "Units Sold", "Total Revenue"],
    });
  };

  return <button onClick={handleDownloadReport}>Download Sales Report</button>;
};

export default ReportComponent;
```

### Vue.js Example

```vue
<template>
  <button @click="downloadReport">Download Report</button>
</template>

<script>
import { ArrayToXLSX } from "ssr-xlsx-export";

export default {
  methods: {
    downloadReport() {
      const data = [
        { name: "Alice", score: 95, grade: "A" },
        { name: "Bob", score: 87, grade: "B" },
        { name: "Charlie", score: 92, grade: "A" },
      ];

      ArrayToXLSX(data, {
        filename: "student-grades",
        sheetName: "Grades",
      });
    },
  },
};
</script>
```

### Array of Arrays Example

```javascript
import { ArrayToXLSX } from "ssr-xlsx-export";

const matrixData = [
  ["Product", "Q1", "Q2", "Q3", "Q4"],
  ["Laptops", 120, 130, 140, 150],
  ["Phones", 200, 220, 210, 230],
  ["Tablets", 80, 90, 85, 95],
];

ArrayToXLSX(matrixData, {
  filename: "quarterly-sales",
  includeHeaders: false, // First row is already headers
});
```

## API Reference

### `ArrayToXLSX(data, options)`

Main function to convert array data to downloadable XLSX file.

#### Parameters

- `data: ArrayData` - Array of objects or array of arrays
- `options: DownloadOptions` - Configuration options (optional)

#### Options

```typescript
interface DownloadOptions {
  filename?: string; // Default: 'report'
  sheetName?: string; // Default: 'Sheet1'
  headers?: string[]; // Custom column headers
  includeHeaders?: boolean; // Default: true
}
```

### Utility Functions

#### `validateArrayData(data: any): boolean`

Validates if the provided data is in the correct format for conversion.

```javascript
import { validateArrayData } from "ssr-xlsx-export";

const data = [{ name: "John", age: 30 }];
if (validateArrayData(data)) {
  console.log("Data is valid for XLSX conversion");
}
```

#### `csvToArray(csvString: string, delimiter?: string): string[][]`

Converts a CSV string to array format suitable for XLSX conversion.

```javascript
import { csvToArray, ArrayToXLSX } from "ssr-xlsx-export";

const csvData = `Name,Age,City
John Doe,30,New York
Jane Smith,25,Los Angeles`;

const arrayData = csvToArray(csvData);
ArrayToXLSX(arrayData, {
  filename: "csv-converted-data",
});
```

## Data Formats Supported

### Array of Objects

```javascript
[
  { name: "John", age: 30, city: "NYC" },
  { name: "Jane", age: 25, city: "LA" },
];
```

### Array of Arrays

```javascript
[
  ["Name", "Age", "City"],
  ["John", 30, "NYC"],
  ["Jane", 25, "LA"],
];
```

## Error Handling

The library includes comprehensive error handling:

```javascript
try {
  ArrayToXLSX(data, options);
  console.log("Download initiated successfully!");
} catch (error) {
  console.error("Download failed:", error.message);

  // Handle specific error cases
  if (error.message.includes("browser environment")) {
    alert("This feature only works in web browsers");
  }
}
```

Common errors:

- Empty or invalid data array
- Browser environment not detected (when used in Node.js)
- XLSX generation failures
- Invalid data format

## Browser Compatibility

This library works in all modern browsers that support:

- **Blob API** - For creating file data
- **URL.createObjectURL** - For creating download URLs
- **File download functionality** - For triggering downloads
- **ES6+ JavaScript features**

### Supported Browsers

- Chrome 52+
- Firefox 52+
- Safari 10+
- Edge 79+
- Opera 39+

## Environment Requirements

- **Browser Environment**: Must run in a web browser
- **JavaScript**: ES6+ support required
- **Module System**: ES modules or CommonJS

## Dependencies

- `xlsx`: For Excel file generation and manipulation

## Installation & Setup

```bash
# Install the package
npm install ssr-xlsx-export

# For TypeScript projects, types are included
# No need for @types/ssr-xlsx-export
```

## Troubleshooting

### "Download functionality is only available in browser environments"

This error occurs when trying to use the library in a Node.js environment. The library is designed for frontend use only.

**Solution**: Use this library only in client-side code (React components, Vue components, vanilla JS in browsers).

### Download not working in some browsers

Ensure your browser supports the Blob API and file downloads. Some corporate firewalls or browser extensions might block automatic downloads.

### Large files causing memory issues

For very large datasets (10,000+ rows), consider:

- Processing data in chunks
- Using streaming solutions for server-side processing
- Implementing pagination for large reports

## Development

```bash
# Clone the repository
git clone https://github.com/SriRakeshKumar/ssr-xlsx-export.git

# Install dependencies
npm install

# Build the library
npm run build

# Run tests
npm test
```

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Create a Pull Request

## Support

If you encounter any issues or have questions:

1. Check the [Browser Compatibility](#browser-compatibility) section
2. Ensure you're using the library in a browser environment
3. File an issue on the GitHub repository with:
   - Browser version
   - Error message
   - Sample code that reproduces the issue

---

**Note**: This is a client-side library designed for browser environments. For server-side Excel generation in Node.js, consider using libraries like `exceljs` or `xlsx` directly with file system operations.
