# PowerPoint-To-TXT

Extracts text from `.pptx` and `.ppt` presentations, including slide content, tables, grouped shapes and speaker notes, then consolidates everything into a structured `.txt` report with logging and error tracking.

## Overview

PowerPoint-To-TXT extracts text from PowerPoint presentations and combines the results into a single `.txt` file. It scans an input directory for `.pptx` and `.ppt` files, processes each presentation slide by slide and captures text from shapes, tables, grouped shapes and speaker notes. The output includes a clear summary of how many files were processed, how many succeeded or failed and the total number of slides extracted. Followed by organized sections for each source.

The script supports both hardcoded defaults and optional arguments for the input directory, output file, log file and recursive scanning. Legacy `.ppt` files are automatically converted to `.pptx` using `LibreOffice` when available. While `.pptx` files are read directly with `python-pptx`. This application also writes detailed logs to help troubleshoot failed extractions, making it useful for archiving, searching, auditing or repurposing text content from large batches of input files.

## Set Up

Below are instructions for installing and running this application on a Linux machine.

### Programs Needed

- [Git](https://git-scm.com/downloads)

- [Python](https://www.python.org/downloads/)

### Steps

1. Install the above programs

2. Open a terminal

3. Clone this repository: `git clone git@github.com:devbret/powerpoint-to-txt.git`

4. Navigate to the repo's directory: `cd powerpoint-to-txt`

5. Create a virtual environment: `python3 -m venv venv`

6. Activate your virtual environment: `source venv/bin/activate`

7. Install the needed dependencies for running the script: `pip install -r requirements.txt`

8. Place your `.pptx` and `.ppt` files in the `input` directory of this repo

9. Use the following command to process: `python3 app.py`

10. The results will be returned to you at the root of this repo as a `.txt` file

11. Exit the virtual environment: `deactivate`

## Other Considerations

This project repo is intended to demonstrate an ability to do the following:

- Extract readable text from `.ppt` and `.pptx` PowerPoint files in a specified input directory

- Convert legacy `.ppt` files into `.pptx` format using `LibreOffice` so their text can be processed

- Write all extracted PowerPoint text into one combined `.txt` file while also logging progress, successes, failures and summary statistics

If you have any questions or would like to collaborate, please reach out either on GitHub or via [my website](https://bretbernhoft.com/).
