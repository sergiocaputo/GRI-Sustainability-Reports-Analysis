# GRI-Sustainability-Reports-Analysis

This project's goal is about the analysis of Sustainability reports which are based on GRI standard disclosures.

Given some PDF reports is possible to extract a textual representation, by OCR, Text-Extractor or Mixed approach, useful to detect the presence of GRI disclosures and to study the polarity of the context in which they are located.

## Installation

    pip install -r requirements.txt

## Usage

Given Sustainability report files in PDF format, place them in the `reports` folder and run the main:

    python .\src\main.py   
    
Different kind of analysis are performed on the same PDF file depending on the approach used to extract the text (OCR, Text-Extractor or Mixed approach).

Information about GRI disclosures presence and their context polarity are available in JSON files that can be used to perform analysis.

Results are compared to the Ground Truth values available in the `210726 Template GRI report analysis.xlsx` file and stored in dedicated Excel files.

## Documentation
