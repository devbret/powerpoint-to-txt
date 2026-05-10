# PowerPoint-To-TXT

Extracts text from `.pptx` and `.ppt` presentations, including slide content, tables, grouped shapes and speaker notes, then consolidates everything into a structured `.txt` report with logging and error tracking.

## Overview

PowerPoint-To-TXT extracts text from PowerPoint presentations and combines the results into a single `.txt` file. It scans an input directory for `.pptx` and `.ppt` files, processes each presentation slide by slide and captures text from shapes, tables, grouped shapes and speaker notes. The output includes a clear summary of how many files were processed, how many succeeded or failed and the total number of slides extracted. Followed by organized sections for each source.

The script supports both hardcoded defaults and optional arguments for the input directory, output file, log file and recursive scanning. Legacy `.ppt` files are automatically converted to `.pptx` using `LibreOffice` when available. While `.pptx` files are read directly with `python-pptx`. This application also writes detailed logs to help troubleshoot failed extractions, making it useful for archiving, searching, auditing or repurposing text content from large batches of input files.
