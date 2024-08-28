package com.example.dlautomation.logic.readers;

import com.example.dlautomation.logic.logging.GlobalLogger;
import com.example.dlautomation.logic.models.ChangeInfo;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

public class DocWordReader extends AbstractWordReader {

    private static final Logger logger = GlobalLogger.getLogger();

    public DocWordReader(String docPath) {
        super(docPath);
        logger.log(Level.INFO, "Initialized DocWordReader for document: {0}", docPath);
    }

    @Override
    public String extractTableName() throws IOException {
        logger.log(Level.INFO, "Attempting to extract table name from document: {0}", docPath);

        if (isTemporaryFile(docPath)) {
            logger.log(Level.WARNING, "Document {0} is a temporary file or not a valid Word document.", docPath);
            throw new IOException("File is a temporary document or not a valid Word file.");
        }

        try (FileInputStream fis = new FileInputStream(docPath)) {
            HWPFDocument document;
            try {
                document = new HWPFDocument(fis);
            } catch (IllegalArgumentException e) {
                logger.log(Level.SEVERE, "Failed to load document: {0}. Error: {1}", new Object[]{docPath, e.getMessage()});
                throw new IOException("The document could not be processed. It might be corrupted or in an unsupported format.", e);
            }

            Range range = document.getRange();
            int numParagraphs = range.numParagraphs();
            logger.log(Level.INFO, "Document contains {0} paragraphs", numParagraphs);

            for (int i = 0; i < numParagraphs; i++) {
                Paragraph paragraph = range.getParagraph(i);
                String paragraphText = paragraph.text();

                if (paragraphText.contains("Wenn")) {
                    String tableName = extractTableNameFromText(paragraphText);
                    if (tableName != null) {
                        logger.log(Level.INFO, "Table name found: {0}", tableName);
                        return tableName;
                    }
                }
            }
        } catch (IOException e) {
            logger.log(Level.SEVERE, "Failed to process document: {0}. Error: {1}", new Object[]{docPath, e.getMessage()});
            throw e;
        } catch (Exception e) {
            logger.log(Level.SEVERE, "Unexpected error while processing document: {0}. Error: {1}", new Object[]{docPath, e.getMessage()});
            throw new IOException("Failed to extract table name", e);
        }
        logger.log(Level.WARNING, "Table name not found in document: {0}", docPath);
        return "Unknown Table Name";
    }

    private boolean isTemporaryFile(String filePath) {
        boolean isTempFile = filePath.startsWith("~$");
        if (isTempFile) {
            logger.log(Level.INFO, "File {0} is identified as a temporary file.", filePath);
        }
        return isTempFile;
    }

    private String extractTableNameFromText(String text) {
        String prefix = "Wenn";
        int startIndex = text.indexOf(prefix);

        if (startIndex == -1) {
            logger.log(Level.FINE, "Prefix 'Wenn' not found in text.");
            return null;
        }

        startIndex += prefix.length();
        int endIndex = text.indexOf('.', startIndex);

        if (endIndex == -1) {
            endIndex = text.length();
        }

        String tableName = text.substring(startIndex, endIndex).trim();
        logger.log(Level.INFO, "Extracted table name: {0}", tableName);
        return tableName;
    }

    @Override
    public String extractReleasestand() throws IOException {
        logger.log(Level.INFO, "Extracting releasestand from document: {0}", docPath);
        String releasestand = "";
        try (FileInputStream fis = new FileInputStream(docPath);
             HWPFDocument document = new HWPFDocument(fis)) {
            Range range = document.getRange();
            String firstPageText = range.text();
            String regex = "Stand:\\s*([^,]+),";
            java.util.regex.Pattern pattern = java.util.regex.Pattern.compile(regex);
            java.util.regex.Matcher matcher = pattern.matcher(firstPageText);

            if (matcher.find()) {
                releasestand = matcher.group(1).trim();
                logger.log(Level.INFO, "Extracted releasestand: {0}", releasestand);
            } else {
                logger.log(Level.WARNING, "Releasestand not found in document: {0}", docPath);
            }
        } catch (IOException e) {
            logger.log(Level.SEVERE, "Failed to extract releasestand from document: {0}. Error: {1}", new Object[]{docPath, e.getMessage()});
            throw e;
        }
        return releasestand;
    }

    @Override
    public List<ChangeInfo> getRedChanges(String tableName, String releasestand) throws IOException {
        logger.log(Level.INFO, "Extracting red changes from document: {0} for table: {1}", new Object[]{docPath, tableName});
        List<ChangeInfo> changes = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(docPath);
             HWPFDocument document = new HWPFDocument(fis)) {
            Range range = document.getRange();
            TableIterator tableIterator = new TableIterator(range);
            while (tableIterator.hasNext()) {
                Table table = tableIterator.next();
                for (int i = 0; i < table.numRows(); i++) {
                    TableRow row = table.getRow(i);
                    if (row.numCells() > 1) {
                        TableCell numberCell = row.getCell(0);
                        TableCell changeCell = row.getCell(1);

                        String changeNumber = numberCell.text().trim();
                        String wholeString = getWholeTextFromRow(row);
                        String changeText = getRedTextFromCell(changeCell);

                        boolean isFullyRed = changeText.equals(changeCell.text().trim());
                        String logik = determineLogik(changeCell);

                        if (!changeText.isEmpty()) {
                            logger.log(Level.INFO, "Red change found: Change Number: {0}, Change Text: {1}, Is Fully Red: {2}", new Object[]{changeNumber, changeText, isFullyRed});
                            changes.add(new ChangeInfo(tableName, changeNumber, changeText, releasestand, getMappingName(), isFullyRed, logik, wholeString));
                        }
                    }
                }
            }
        } catch (IOException e) {
            logger.log(Level.SEVERE, "Failed to extract red changes from document: {0}. Error: {1}", new Object[]{docPath, e.getMessage()});
            throw e;
        }
        return changes;
    }

    private String determineLogik(TableCell cell) {
        StringBuilder cellText = new StringBuilder();
        boolean isCrossedOut = false;

        for (int i = 0; i < cell.numCharacterRuns(); i++) {
            CharacterRun run = cell.getCharacterRun(i);
            String color = Integer.toHexString(run.getColor());
            if ("FF0000".equalsIgnoreCase(color)) {
                cellText.append(run.text().trim()).append(" ");
                if (run.isStrikeThrough()) {
                    isCrossedOut = true;
                }
            }
        }

        String logik = isCrossedOut ? "Rückbau Logik" : "Neue Logik";
        logger.log(Level.INFO, "Determined logic: {0}", logik);
        return logik;
    }

    private String getRedTextFromCell(TableCell cell) {
        StringBuilder redText = new StringBuilder();
        boolean hasRedText = false;

        for (int i = 0; i < cell.numCharacterRuns(); i++) {
            CharacterRun run = cell.getCharacterRun(i);
            String text = run.text().trim();
            int colorIndex = run.getColor();

            if (isRedColor(colorIndex)) {
                redText.append(text).append(" ");
                hasRedText = true;
            }
        }

        if (hasRedText) {
            logger.log(Level.INFO, "Red text found in cell: {0}", redText.toString().trim());
        } else {
            logger.log(Level.FINE, "No red text found in cell.");
        }

        return redText.toString().trim();
    }

    private boolean isRedColor(int colorIndex) {
        return colorIndex == 6;
    }

    private String getWholeTextFromRow(TableRow row) {
        StringBuilder wholeText = new StringBuilder();
        for (int i = 0; i < row.numCells(); i++) {
            TableCell cell = row.getCell(i);
            if (containsRedText(cell)) {
                wholeText.append(getWholeText(cell)).append(" | ");
            }
        }
        logger.log(Level.INFO, "Extracted whole text from row: {0}", wholeText.toString().trim());
        return wholeText.toString().trim();
    }

    private boolean containsRedText(TableCell cell) {
        for (int i = 0; i < cell.numCharacterRuns(); i++) {
            CharacterRun run = cell.getCharacterRun(i);
            if (isRedColor(run.getColor())) {
                return true;
            }
        }
        return false;
    }

    private String getWholeText(TableCell cell) {
        StringBuilder cellText = new StringBuilder();
        for (int i = 0; i < cell.numCharacterRuns(); i++) {
            CharacterRun run = cell.getCharacterRun(i);
            cellText.append(run.text().trim());
        }
        return cellText.toString().trim();
    }
}
