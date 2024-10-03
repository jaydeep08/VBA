// Function to extract the file name from a URL with extension verification for Microsoft environment
const extractFileNameFromLink = (url) => {
  try {
    // Create a new URL object
    const link = new URL(url);

    // Get the pathname from the URL object and split it to find the last part (file name with possible query params)
    let fileName = link.pathname.split('/').pop();

    // Check if the fileName contains a query string (?)
    if (fileName.includes('?')) {
      // Strip out the query string
      fileName = fileName.split('?')[0];
    }

    // Verify the file has a valid extension
    const validExtensions = [
      'doc', 'docx', 'rtf', 'txt', 'pdf', 'odt', 'dot', 'dotx',  // Documents
      'xls', 'xlsx', 'csv', 'ods', 'xlt', 'xltx',                 // Spreadsheets
      'ppt', 'pptx', 'pps', 'ppsx', 'potx',                       // Presentations
      'jpg', 'jpeg', 'png', 'gif', 'bmp', 'tif', 'tiff',          // Images
      'mp3', 'wav', 'mp4', 'wmv', 'avi', 'mov', 'mkv',            // Audio/Video
      'zip', 'rar', '7z', 'tar', 'gz',                            // Compressed Files
      'xml', 'json', 'html', 'xhtml', 'css', 'js',                // Web/Code Files
      'psd', 'ai', 'svg',                                         // Design Files
      'ics', 'exe', 'dll', 'iso'                                  // Miscellaneous
    ];

    const fileExtension = fileName.split('.').pop().toLowerCase();

    // Check if the file name has one of the valid extensions
    if (validExtensions.includes(fileExtension)) {
      return fileName;
    } else {
      return ''; // Return empty if no valid extension is found
    }
  } catch (error) {
    // In case of an invalid URL or other errors, return an empty string
    return '';
  }
};

// Example usage:
const fileLinks = [
  "https://example.com/docs/report.docx",
  "https://example.com/files/presentation.pptx?version=2",
  "https://example.com/files/spreadsheet.xlsx",
  "https://example.com/downloads/image.jpg",
  "https://example.com/files/no-extension",
  "https://example.com/files/archive.zip?download=true",
  "invalidLink"
];

// Extract file names from the links
const fileNames = fileLinks.map(link => extractFileNameFromLink(link));

console.log(fileNames); // Output: ["report.docx", "presentation.pptx", "spreadsheet.xlsx", "image.jpg", "", "archive.zip", ""]
