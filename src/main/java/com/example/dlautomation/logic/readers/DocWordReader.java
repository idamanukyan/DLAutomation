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
        try (FileInputStream fis = new FileInputStream(docPath);
             HWPFDocument document = new HWPFDocument(fis)) {

            Range range = document.getRange();
            int numParagraphs = range.numParagraphs();
            String tableName = "Unknown Table Name";

            for (int i = 0; i < numParagraphs; i++) {
                Paragraph paragraph = range.getParagraph(i);
                String paragraphText = cleanHyperlinks(paragraph.text().trim());

                if (i + 1 < numParagraphs) {
                    Paragraph nextParagraph = range.getParagraph(i + 1);
                    if (isTableStart(nextParagraph)) {
                        tableName = paragraphText;
                        break;
                    }
                }
            }

            if (tableName.equals("Unknown Table Name")) {
                for (int i = 0; i < numParagraphs; i++) {
                    Paragraph paragraph = range.getParagraph(i);
                    String paragraphText = cleanHyperlinks(paragraph.text().trim());
                    if (!paragraphText.isEmpty() && i + 1 < numParagraphs) {
                        Paragraph nextParagraph = range.getParagraph(i + 1);
                        if (isTableStart(nextParagraph)) {
                            tableName = paragraphText;
                            break;
                        }
                    }
                }
            }

            return tableName;
        }
    }

    private boolean isTableStart(Paragraph paragraph) {
        return paragraph.text().trim().isEmpty();
    }

    private String cleanHyperlinks(String text) {
        String cleanedText = text;

        cleanedText = cleanedText.replaceAll("https?://\\S+", "");
        cleanedText = cleanedText.replaceAll("\\[HYPERLINK[^\\]]*\\]", "");
        cleanedText = cleanedText.replaceAll("\\p{C}", "");
        cleanedText = cleanedText.trim();

        return cleanedText;
    }

    @Override
    public String extractReleasestand() throws IOException {
        try (FileInputStream fis = new FileInputStream(docPath);
             HWPFDocument document = new HWPFDocument(fis)) {
            Range range = document.getRange();
            TableIterator tableIterator = new TableIterator(range);

            while (tableIterator.hasNext()) {
                Table table = tableIterator.next();
                for (int i = 0; i < table.numRows(); i++) {
                    TableRow row = table.getRow(i);
                    if (row.numCells() > 1) {
                        TableCell cell = row.getCell(0);
                        String cellText = cleanHyperlinks(cell.text().trim());
                        if (cellText.contains("Releasestand")) {
                            TableCell releasestandCell = row.getCell(1);
                            return cleanHyperlinks(releasestandCell.text().trim());
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
                        String changeNumber = cleanHyperlinks(numberCell.text().trim());
                        String changeText = cleanHyperlinks(changeCell.text().trim());
                        String wholeText = getRedTextFromCell(changeCell);

                        boolean isFullyRed = isTextFullyRed(changeCell);
                        String logik = determineLogik(changeCell);

                        if (!changeText.isEmpty()) {
                            changes.add(new ChangeInfo(tableName, changeNumber, changeText, releasestand, getMappingName(), isFullyRed, logik, wholeText));
                        }
                    }
                }
            }
        }
        return changes;
    }

    private String determineLogik(TableCell cell) {
        String cellText = cleanHyperlinks(cell.text().trim());
        if (cellText.contains("deleted")) {
            return "Rückbau Logik";
        }
        return "Neuer Variable";
    }

    private boolean isTextFullyRed(TableCell cell) {
        // Placeholder for actual implementation
        return cell.text().contains("red text indicator");
    }

    private String getRedTextFromCell(TableCell cell) {
        boolean hasRedText = false;
        StringBuilder cellText = new StringBuilder();
       // for (CharacterRun run : cell.getCharacterRuns()) {
         //   if (run.getIcofo() == 0xFF0000) { // Check for red color
                hasRedText = true;
           // }
           // cellText.append(run.text());
       // }

        // Only return the text if there is at least one red letter
     //   return hasRedText ? cellText.toString().trim() : "";

        return null;
    }
}
