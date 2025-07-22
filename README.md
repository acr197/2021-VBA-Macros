# Excel VBA Automation: Placement & Creative QA Toolkit
Refreshing my 2021 VBA macro that I built pre-GPT that benefitted greatly from an AI refresh. Always looking to improve legacy processes 👍

## Overview  
This toolkit is a rebuilt compilation of simple tools (at least the macros without sensitive details specialized for clients) that I made at home that was used for many QA and reporting tasks across media operations.

This sheet builds on VBA I originally learned through StackOverflow, MrExcel, and VBAExpress. With that foundation and GenAI, I rebuilt a cleaner, more dynamic version in under an hour - something that used to take weeks manually to learn, build, test, and deploy now took me under an hour with my knowledge of coding logic and prompt engineering.

This is one of my favorite examples of AI-assisted coding: a simple VBA tool that saves teams hours on repetitive, tedious tasks while improving standardization across files, QA, and naming. It's a toolkit thats not about replacing people - it’s about delegating the annoying parts to the robots so humans can focus on strategy and insights.

## Tech Used  
- Excel VBA (self‑taught pre‑LLM, later AI‑tuned)  
- FileDialog integration  
- Dynamic header detection  
- Conditional formatting  
- Data validation  

## How It Works  
Each macro batch runs via the VBA editor or toolbar buttons to:  
- Reset Excel defaults after errors  
- Parse taxonomy strings into structured columns  
- Merge selected workbooks into a master sheet  
- Format CSV or SQL exports with styles and freeze pane  
- Save workbooks to Downloads with timestamped names  
- Unlock Protected View files for other macros  
- Remove placeholder tildes in data ranges  

## Highlighted Functions  

### 1. ResetSettings  
Restores alerts, screen updating, calculation mode, and scrollbars to troubleshoot display or behavior issues.

### 2. TaxonomyKeyParse  
Parses client taxonomy keys (KEY~Value_) into separate columns for each key.

### 3. MergeExcelSheets  
Prompts file selection and consolidates data from multiple workbooks into one sheet without duplicate headers.

### 4. FormatCSV  
Styles header row with bold fill and center alignment, auto‑fits columns, and freezes the top pane.

### 5. QuickSave  
Saves the active workbook as XLSX in the Downloads folder with a timestamped name or prompts manual save on error.

### 6. UnprotectSheets  
Opens Protected View workbooks in edit mode so automated routines can run without interruption.

### 7. ReplaceTildes  
Replaces “~~” in selected ranges to simplify super-common lookup placeholders, ctrl-f, filter searches, and clean up data.

## Business Value & Use Cases  
- **Time savings**: weeks of manual coding into hours with AI assistance  
- **Error reduction**: prevents silent failures and enforces standards  
- **Team efficiency**: non‑technical users run macros via buttons  
- **Scalability**: handles varying data sizes with no hardcoded ranges  
- **Human augmentation**: frees employees to focus on strategic work and innovation  

## Notes  
- Built manually before LLMs and later refined with AI assistants  
- Saves the team many hours of tedious work each month  
- Demonstrates AI as an assistant to augment human capabilities, not replace jobs
- This is not the original code used at my place of employment as the company owns that. This was re-written from scratch as a demo of GenAI on my personal time to display in a compliant manner.
