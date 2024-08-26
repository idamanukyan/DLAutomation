package com.example.dlautomation.logic.readers;

import com.example.dlautomation.logic.models.ChangeInfo;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class DocWordReader extends AbstractWordReader {

    public DocWordReader(String docPath) {
        super(docPath);
    }

    @Override
    public String extractTableName() throws IOException {

        if (isTemporaryFile(docPath)) {
            throw new IOException("File is a temporary document or not a valid Word file.");
        }

        try (FileInputStream fis = new FileInputStream(docPath)) {
            HWPFDocument document;
            try {
                document = new HWPFDocument(fis);
            } catch (IllegalArgumentException e) {
                System.err.println("Failed to load document: " + docPath);
                throw new IOException("The document could not be processed. It might be corrupted or in an unsupported format.", e);
            }

            Range range = document.getRange();
            int numParagraphs = range.numParagraphs();

            for (int i = 0; i < numParagraphs; i++) {
                Paragraph paragraph = range.getParagraph(i);
                String paragraphText = paragraph.text();

                if (paragraphText.contains("Wenn")) {
                    String tableName = extractTableNameFromText(paragraphText);
                    if (tableName != null) {
                        System.out.println("Table name found: " + tableName);
                        return tableName;
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("Failed to load document: " + docPath + ". Error: " + e.getMessage());
            throw e;
        } catch (Exception e) {
            System.err.println("Unexpected error processing document: " + docPath + ". Error: " + e.getMessage());
            throw new IOException("Failed to extract table name", e);
        }
        return "Unknown Table Name";
    }

    private boolean isTemporaryFile(String filePath) {
        return filePath.startsWith("~$");
    }


    private String extractTableNameFromText(String text) {
        String prefix = "Wenn";
        int startIndex = text.indexOf(prefix);

        if (startIndex == -1) {
            return null;
        }

        startIndex += prefix.length();
        int endIndex = text.indexOf('.', startIndex);

        if (endIndex == -1) {
            endIndex = text.length();
        }

        return text.substring(startIndex, endIndex).trim();
    }

    @Override
    public String extractReleasestand() throws IOException {
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
            }
        }
        return releasestand;
    }

    @Override
    public List<ChangeInfo> getRedChanges(String tableName, String releasestand) throws IOException {
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
                            changes.add(new ChangeInfo(tableName, changeNumber, changeText, releasestand, getMappingName(), isFullyRed, logik, wholeString));
                        }
                    }
                }
            }
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

        return isCrossedOut ? "Rückbau Logik" : "Neuer Variable";
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
        return hasRedText ? redText.toString().trim() : "";
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
