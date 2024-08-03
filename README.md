# Troubleshooting Aspose.NET Document Generation

This repository aims to investigate and troubleshoot issues related to DOCX file generation using Aspose.NET in a Dynamics 365 F&O implementation. Specifically, we address a problem where a generated DOCX file becomes corrupted when it reaches four pages.

## Repository Structure

- **1x working.docx**: A sample DOCX file with one page of content.
- **2x working.docx**: A sample DOCX file with two pages of content.
- **3x working.docx**: A sample DOCX file with three pages of content.
- **4x corrupted.docx**: A sample DOCX file with four pages of content, which is corrupted and cannot be opened.
- **README.md**: This file.
- **analysis**: Directory containing scripts and tools for analyzing DOCX files.

## Problem Description

The DOCX files generated using Aspose.NET work correctly for one, two, and three pages. However, when the document reaches four pages, it becomes corrupted and cannot be opened in standard DOCX readers like Microsoft Word.

## Steps to Reproduce

1. **Environment Setup**:
   - Ensure you are using Dynamics 365 F&O with Correspondentie versie: 3.1.
   - Use SGdoc.dll assembly with version: server: 3.0.5533.30207 / client: 3.0.5533.30207.

2. **Prepare the Template**:
   - Use an empty template (`sjabloon`) containing only a rectangular shape.

3. **Generate DOCX Files**:
   - In Dynamics 365 F&O, generate DOCX files from a module (similar to contracts) using the prepared template:
     - Generate a DOCX file with content equivalent to 1 page.
     - Generate a DOCX file with content equivalent to 2 pages.
     - Generate a DOCX file with content equivalent to 3 pages.
     - Generate a DOCX file with content equivalent to 4 pages.

4. **Open the Generated DOCX Files**:
   - Open the generated DOCX files in a DOCX reader such as Microsoft Word.
   - Observe that the document with four pages is corrupted and cannot be opened.

## Analysis

To identify the root cause of the corruption, follow these steps:

1. **Extract XML Content**:
   - Extract the XML files from the DOCX packages for 1, 2, 3, and 4 pages.
2. **Compare Structures**:
   - Analyze the XML structures of the working DOCX files (1, 2, and 3 pages) and the corrupted DOCX file (4 pages).
3. **Identify Inconsistencies**:
   - Look for missing elements, incorrect namespaces, and any discrepancies in the XML structures.
4. **Tools Used**:
   - Python scripts are provided in the `analysis` directory to automate the extraction and comparison of XML contents.

## Scripts

The `analysis` directory contains the following scripts:

- **extract_xml.py**: Script to extract XML files from DOCX packages.
- **compare_xml.py**: Script to compare XML structures and identify inconsistencies.

## Conclusion

The goal is to pinpoint the exact cause of the corruption in the four-page DOCX file by comparing it against the working documents. By identifying inconsistencies or missing elements, we can propose a solution to fix the generation process in Aspose.NET.

## Contribution

Contributions are welcome! If you have insights or potential fixes, please open an issue or submit a pull request.

## License

This project is licensed under the MIT License.