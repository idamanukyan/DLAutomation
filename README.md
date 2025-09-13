📄 DocxWordReader – Document Change Extraction Tool

A Java-based tool for reading and analyzing Microsoft Word .docx documents.
It extracts metadata (like table name and release version) and identifies changes highlighted in red text, converting them into structured data objects for further processing.

🎯 Overview

This component is part of the DL Automation project.
Its main responsibilities include:

Parsing .docx Word files using Apache POI

Extracting metadata such as table names and release versions (Releasestand)

Detecting red-colored changes in tables and classifying them as Neue Logik (new logic) or Rückbau Logik (rollback logic)

Logging events with a centralized GlobalLogger for auditing and debugging

✨ Key Features

📑 Table Name Extraction
Reads the table cell labeled "Tabellenname/View" and returns the corresponding value.

📅 Release Version Extraction
Finds and retrieves the "Releasestand" field from the document.

📝 Change Tracking

Scans table rows for red-colored text.

Identifies both full-cell red changes and partial highlights.

Captures change number, change text, logic type (Neue Logik or Rückbau Logik), and full raw text.

🔍 Logic Determination

Red text with strikethrough → Rückbau Logik

Red text without strikethrough → Neue Logik

📊 Structured Output
Returns results as ChangeInfo objects containing:

Table Name

Change Number

Change Text (red-only)

Releasestand

Mapping Name

Boolean: fully red or mixed

Logic classification

Full raw text

⚠️ Error Handling

Skips temporary files (~$filename.docx)

Gracefully handles missing metadata (returns "Unknown Table Name" or "Unknown Releasestand")

Logs warnings and errors via GlobalLogger

🛠️ Tech Stack

Language: Java

Library: Apache POI (org.apache.poi.xwpf.usermodel)

Logging: Java Util Logging (via custom GlobalLogger)

Architecture: Extends abstract base class AbstractWordReader

📂 Project Structure

DocxWordReader – Core class for .docx parsing

AbstractWordReader – Base abstraction for different Word readers (.doc vs .docx)

ChangeInfo – DTO for holding extracted change details

GlobalLogger – Centralized logging utility

🚀 Usage Example
DocxWordReader reader = new DocxWordReader("document.docx");

String tableName = reader.extractTableName();
String releasestand = reader.extractReleasestand();

List<ChangeInfo> changes = reader.getRedChanges(tableName, releasestand);

for (ChangeInfo change : changes) {
    System.out.println(change);
}

📊 Example Output
Table: RMS_Table_01
Releasestand: V2.1

Change Number: 101
Change Text: Added new validation rule
Logic: Neue Logik
Fully Red: false
Whole Text: Added new validation rule (mandatory for all records)

Change Number: 102
Change Text: Deprecated field XYZ
Logic: Rückbau Logik
Fully Red: true
Whole Text: Deprecated field XYZ

📜 License

This project is licensed under the MIT License – see LICENSE
 for details.
