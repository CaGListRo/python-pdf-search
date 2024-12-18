# PDF Search and Export Tool

This Python script allows you to search for a specific term within PDF files in a specified folder. The search results, including the page number and surrounding text context, are exported to a `.docx` file for easy review.

## Features

- **Search PDFs**: Scans all PDF files in a given folder for a specific search term.
- **Text Context**: Provides the surrounding text around the found search term.
- **Highlighting**: Highlights the search term in the exported Word document.
- **Output Format**: Saves results in a structured `.docx` file, organized by file and page.

## Example

Included with this tool is a test PDF, "The Great Toaster Escape," that can be used to test the functionality.

## Prerequisites

Ensure you have the following Python libraries installed:

- [PyMuPDF](https://pypi.org/project/PyMuPDF) (fitz)
- [python-docx](https://pypi.org/project/python-docx)

Install them using:

```bash
pip install pymupdf python-docx
```

## Usage

### Script Execution

1. **Clone the Script**: Copy the script to your local environment.
2. **Run the Script**: Execute the script in a terminal or command prompt.
3. **Inputs**:
   - **Folder Path**: Enter the folder path containing PDF files.
   - **Search Term**: Provide the term you want to search for.
   - **Output Path**: Specify the directory where the results should be saved.

```bash
python script_name.py
```

### Output

The results are saved in a `.docx` file named `occurrence_of_<search_term>.docx` in the specified output directory. The document includes:

- File names of the PDFs where the term was found.
- Page numbers of occurrences.
- Context text with the search term highlighted in bold red.

## Code Overview

### Functions

1. **`search_in_pdf(file_path, search_term)`**

   - Scans a single PDF for the search term.
   - Returns a list of dictionaries containing page numbers, text context, and term index.

2. **`search_in_folder(folder_path, search_term)`**

   - Iterates through all PDFs in a folder.
   - Calls `search_in_pdf` for each file and organizes results by file name.

3. **`highlight_search_term(paragraph, search_term)`**

   - Adds the search term in bold red font to the `.docx` document.

4. **`sanitize_text(text)`**

   - Cleans non-printable characters, preserving line breaks and tabs.

5. **`write_results_to_docx(all_results, search_term, output_path)`**
   - Writes search results into a `.docx` file.

### Example Workflow

Given a folder containing the test PDF "The Great Toaster Escape":

1. The script scans the file for your specified term.
2. All matches, including the surrounding context, are written to a `.docx` file.

### Error Handling

- **Invalid Folder Path**: Prompts for a valid directory.
- **No PDFs Found**: Notifies the user if no PDFs exist in the folder.
- **File Permission Issues**: Ensures the output directory is writable.

## Sample Run

```bash
Enter the folder path containing PDF files: D:\PDFs
Enter the search term: toaster
Enter the output DOCX file path: D:\Results
```

Output file: `D:\Results\occurrence_of_toaster.docx`

## Limitations

- The search term is case-insensitive.
- Text extraction accuracy depends on the quality and format of the PDF (e.g., scanned images are not supported).

## License

This script is open-source and free to use.

---

Enjoy exploring the "The Great Toaster Escape" or searching your own documents!
