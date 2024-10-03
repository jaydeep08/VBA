// Function to extract the file name from SharePoint/OneDrive links, including cases with 'file=' in query params
const extractFileNameFromLink = (url) => {
  try {
    // Create a new URL object
    const link = new URL(url);

    // Get the pathname from the URL object (the part before ?)
    let fileName = link.pathname.split('/').pop();

    // If there's a query string, check if 'file=' exists in the query
    if (link.search) {
      // Extract the query part after ?
      const queryParams = link.search.substring(1); // Remove the "?"
      
      // Split query parameters to see if the file name is embedded in 'file='
      const params = new URLSearchParams(queryParams);
      
      // Check if 'file=' is a part of the query parameters
      if (params.has('file')) {
        fileName = params.get('file').split('/').pop(); // Extract the file name from the 'file=' param
      }
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

    // Extract file extension
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
  "https://example.sharepoint.com/docs/report.pdf",                       // Normal path
  "https://example.sharepoint.com/download?file=/documents/document.docx", // File param with file=
  "https://example.sharepoint.com/files/download?file=/spreadsheet.xlsx",  // File param
  "https://example.sharepoint.com/files?file=/images/image.jpg&size=large",// File param with additional query params
  "https://example.sharepoint.com/files/no-extension",                    // No valid extension
  "https://example.sharepoint.com/files/?file=/folder/file.txt&download=true", // File param with folder structure
  "https://example.sharepoint.com/files?somethingrandom"                  // No file name or extension
];

// Extract file names from the links
const fileNames = fileLinks.map(link => extractFileNameFromLink(link));

console.log(fileNames); 
// Output: ["report.pdf", "document.docx", "spreadsheet.xlsx", "image.jpg", "", "file.txt", ""]
