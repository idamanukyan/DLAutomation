package com.example.dlautomation.logic.readers;

import com.example.dlautomation.logic.models.ChangeInfo;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class DocxWordReader extends AbstractWordReader {

    public DocxWordReader(String docPath) {
        super(docPath);
    }

    @Override
    public String extractTableName() throws IOException {
        if (isTemporaryFile(docPath)) {
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
                            return tableNameCell.getText().trim();
                        }
                    }
                }
            }
        } catch (IOException e) {
            // Log the exception with file path for debugging
            System.err.println("Failed to load document: " + docPath + ". Error: " + e.getMessage());
            throw e;
        } catch (Exception e) {
            // Catch any other exceptions that may occur
            System.err.println("Unexpected error processing document: " + docPath + ". Error: " + e.getMessage());
            throw new IOException("Failed to extract table name", e);
        }
        return "Unknown Table Name";
    }

    private boolean isTemporaryFile(String filePath) {
        return filePath.startsWith("~$");
    }


    @Override
    public String extractReleasestand() throws IOException {
        try (FileInputStream fis = new FileInputStream(docPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    if (row.getTableCells().size() > 1) {
                        XWPFTableCell cell = row.getCell(0);
                        String cellText = cell.getText().trim();
                        if (cellText.contains("Releasestand")) {
                            XWPFTableCell releasestandCell = row.getCell(1);
                            return releasestandCell.getText().trim();
                        }
                    }
                }
            }
        }
        return "Unknown Releasestand";
    }

    @Override
    public List<ChangeInfo> getRedChanges(String tableName, String releasestand) throws IOException {
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
                            changes.add(new ChangeInfo(tableName, changeNumber, changeText, releasestand, getMappingName(), isFullyRed, logik, wholeString));
                        }
                    }
                }
            }
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

        if (isCrossedOut) {
            return "Rückbau Logik";
        } else {
            return "Neuer Variable";
        }
    }


    private String getRedTextFromCell(XWPFTableCell cell) {
        StringBuilder redText = new StringBuilder();
        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                if ("FF0000".equalsIgnoreCase(run.getColor())) {
                    redText.append(run.getText(0)).append(" ");
                }
            }
        }
        return redText.toString().trim();
    }

    private String getWholeText(XWPFTableCell cell) {
        boolean hasRedText = false;
        StringBuilder cellText = new StringBuilder();

        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String color = run.getColor();
                if ("FF0000".equalsIgnoreCase(color)) {
                    hasRedText = true;
                }
                cellText.append(run.getText(0));
            }
        }

        return hasRedText ? cellText.toString().trim() : "";
    }

}
