package com.example.dlautomation.logic.readers;

import com.example.dlautomation.logic.logging.GlobalLogger;
import com.example.dlautomation.logic.models.ChangeInfo;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

public class DocxWordReader extends AbstractWordReader {

    private static final Logger logger = GlobalLogger.getLogger();

    public DocxWordReader(String docPath) {
        super(docPath);
        logger.log(Level.INFO, "Initialized DocxWordReader for document: {0}", docPath);
    }

    @Override
    public String extractTableName() throws IOException {
        logger.log(Level.INFO, "Attempting to extract table name from document: {0}", docPath);

        if (isTemporaryFile(docPath)) {
            logger.log(Level.WARNING, "Document {0} is a temporary file or not a valid Word document.", docPath);
            throw new IOException("File is a temporary document or not a valid Word file.");
        }

        try (FileInputStream fis = new FileInputStream(docPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    if (row.getTableCells().size() > 1) {
                        XWPFTableCell cell = row.getCell(0);
                        String cellText = cell.getText().trim();
                        if (cellText.contains("Tabellenname/View")) {
                            XWPFTableCell tableNameCell = row.getCell(1);
                            String tableName = tableNameCell.getText().trim();
                            logger.log(Level.INFO, "Table name found: {0}", tableName);
                            return tableName;
                        }
                    }
                }
            }
        } catch (IOException e) {
            logger.log(Level.SEVERE, "Failed to load document: {0}. Error: {1}", new Object[]{docPath, e.getMessage()});
            throw e;
        } catch (Exception e) {
            logger.log(Level.SEVERE, "Unexpected error processing document: {0}. Error: {1}", new Object[]{docPath, e.getMessage()});
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

    @Override
    public String extractReleasestand() throws IOException {
        logger.log(Level.INFO, "Extracting releasestand from document: {0}", docPath);

        try (FileInputStream fis = new FileInputStream(docPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    if (row.getTableCells().size() > 1) {
                        XWPFTableCell cell = row.getCell(0);
                        String cellText = cell.getText().trim();
                        if (cellText.contains("Releasestand")) {
                            XWPFTableCell releasestandCell = row.getCell(1);
                            String releasestand = releasestandCell.getText().trim();
                            logger.log(Level.INFO, "Extracted releasestand: {0}", releasestand);
                            return releasestand;
                        }
                    }
                }
            }
        } catch (IOException e) {
            logger.log(Level.SEVERE, "Failed to extract releasestand from document: {0}. Error: {1}", new Object[]{docPath, e.getMessage()});
            throw e;
        }

        logger.log(Level.WARNING, "Releasestand not found in document: {0}", docPath);
        return "Unknown Releasestand";
    }

    @Override
    public List<ChangeInfo> getRedChanges(String tableName, String releasestand) throws IOException {
        logger.log(Level.INFO, "Extracting red changes from document: {0} for table: {1}", new Object[]{docPath, tableName});
        List<ChangeInfo> changes = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(docPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    if (row.getTableCells().size() > 1) {
                        XWPFTableCell numberCell = row.getCell(0);
                        XWPFTableCell changeCell = row.getCell(1);

                        String changeNumber = numberCell.getText().trim();
                        String changeText = getRedTextFromCell(changeCell);
                        String wholeString = getWholeText(changeCell);

                        boolean isFullyRed = getRedTextFromCell(changeCell).equals(changeCell.getText());
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

    private String determineLogik(XWPFTableCell cell) {
        StringBuilder cellText = new StringBuilder();
        boolean isCrossedOut = false;

        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String color = run.getColor();
                if (color != null && "FF0000".equalsIgnoreCase(color)) {
                    String text = run.getText(0);
                    if (text != null) {
                        cellText.append(text).append(" ");
                    }

                    if (run.isStrikeThrough()) {
                        isCrossedOut = true;
                    }
                }
            }
        }

        String logik = isCrossedOut ? "Rückbau Logik" : "Neue Logik";
        logger.log(Level.INFO, "Determined logic: {0}", logik);
        return logik;
    }

    private String getRedTextFromCell(XWPFTableCell cell) {
        StringBuilder redText = new StringBuilder();
        boolean hasRedText = false;

        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                if ("FF0000".equalsIgnoreCase(run.getColor())) {
                    redText.append(run.getText(0)).append(" ");
                    hasRedText = true;
                }
            }
        }

        if (hasRedText) {
            logger.log(Level.INFO, "Red text found in cell: {0}", redText.toString().trim());
        } else {
            logger.log(Level.FINE, "No red text found in cell.");
        }

        return redText.toString().trim();
    }

    private String getWholeText(XWPFTableCell cell) {
        StringBuilder cellText = new StringBuilder();
        boolean hasRedText = false;

        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String color = run.getColor();
                if ("FF0000".equalsIgnoreCase(color)) {
                    hasRedText = true;
                }
                cellText.append(run.getText(0));
            }
        }

        String wholeText = cellText.toString().trim();
        logger.log(Level.INFO, "Extracted whole text from cell: {0}", wholeText);
        return hasRedText ? wholeText : "";
    }
}
