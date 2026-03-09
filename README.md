MATRIXER — BLAST Pairwise Matrix Builder

MATRIXER is a lightweight Python GUI application designed to convert BLAST alignment reports into a fully formatted pairwise comparison matrix in Microsoft Excel. The program parses BLAST output files and automatically constructs a symmetric matrix containing query coverage and DNA identity values between all detected accession numbers.

This tool was developed to simplify the generation of pairwise genomic similarity matrices for comparative genomics studies, particularly when large numbers of BLAST comparisons need to be summarized in a structured format suitable for downstream statistical analysis.

The application reads BLAST text reports, extracts accession identifiers and alignment statistics, and generates a square matrix where each accession appears both as a row and as a column. For each pairwise comparison, the matrix stores query coverage and percent identity values. The resulting Excel file is automatically formatted for readability and immediate use in further analysis.

The program supports BLAST report files in standard text format as well as ZIP archives containing BLAST reports. If an existing matrix file is provided, MATRIXER can update the matrix by filling in missing comparisons while preserving existing values.

The output matrix is structured so that each accession is represented by two rows: one containing query coverage values and the other containing DNA identity values. Self-comparisons are represented by a dash (“-”), while missing alignments are marked as “NSS”. Numeric values less than or equal to 1 are displayed using the symbol “≤1”.

To use the program, the user launches the script and selects a BLAST report file. Optionally, an existing matrix file may also be selected if the user wishes to update a previously generated matrix. The user then selects an output directory and provides a filename for the resulting Excel file. Once the run process is initiated, the program parses the BLAST report, constructs the matrix, formats the Excel output, and saves the completed file to the specified location.

The generated Excel matrix contains merged headers, accession labels, and formatted borders to improve readability. Rows and columns are arranged symmetrically so that each accession can be compared directly against all others. The resulting file is suitable for visualization, statistical analysis, or integration into downstream computational workflows.

Typical workflows using MATRIXER involve performing pairwise genome comparisons using BLAST, exporting the BLAST alignment reports, processing these reports with MATRIXER to generate an identity matrix, and then analyzing the resulting matrix using statistical tools such as R or Python.

The program requires Python version 3.8 or higher. It depends on the Python library openpyxl for Excel file generation, while the graphical user interface is built using tkinter, which is included with most standard Python installations.

This software was developed by Davit Janelidze and Ana Gamkrelidze at the One Health Institute to facilitate rapid generation of BLAST-derived pairwise matrices for comparative genomics and bacteriophage genome research.
