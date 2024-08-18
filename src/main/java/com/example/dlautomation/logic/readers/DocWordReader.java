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

                        boolean isFullyRed = isTextFullyRed(changeCell);
                        String logik = determineLogik(changeCell);

                        if (!changeText.isEmpty()) {
                            changes.add(new ChangeInfo(tableName, changeNumber, changeText, releasestand, getMappingName(), isFullyRed, logik));
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

    public void debugParagraphs() throws IOException {
        try (FileInputStream fis = new FileInputStream(docPath);
             HWPFDocument document = new HWPFDocument(fis)) {

            Range range = document.getRange();
            int numParagraphs = range.numParagraphs();

            for (int i = 0; i < numParagraphs; i++) {
                Paragraph paragraph = range.getParagraph(i);
                String paragraphText = cleanHyperlinks(paragraph.text().trim());

                System.out.println("Paragraph " + i + ": " + paragraphText);

                for (int j = 0; j < paragraph.numCharacterRuns(); j++) {
                    CharacterRun run = paragraph.getCharacterRun(j);
                    String text = run.text();
                    if (text.contains("HYPERLINK")) {
                        System.out.println("Found hyperlink in Paragraph " + i + ": " + text);
                    }
                }
            }
        }
    }
}
