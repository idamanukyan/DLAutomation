package com.example.dlautomation.logic.inheritance;

import com.example.dlautomation.logic.ChangeInfo;
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
            TableIterator tableIterator = new TableIterator(range);

            while (tableIterator.hasNext()) {
                Table table = tableIterator.next();
                for (int i = 0; i < table.numRows(); i++) {
                    TableRow row = table.getRow(i);
                    if (row.numCells() > 1) {
                        TableCell cell = row.getCell(0);
                        String cellText = cell.text().trim();
                        if (cellText.contains("Tabellenname/View")) {
                            TableCell tableNameCell = row.getCell(1);
                            return tableNameCell.text().trim();
                        }
                    }
                }
            }
        }
        return "Unknown Table Name";
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
                        String cellText = cell.text().trim();
                        if (cellText.contains("Releasestand")) {
                            TableCell releasestandCell = row.getCell(1);
                            return releasestandCell.text().trim();
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
                        String changeNumber = numberCell.text().trim();
                        String changeText = changeCell.text().trim();

                        boolean isFullyRed = isTextFullyRed(changeCell);
                        String logik = determineLogik(changeCell); // Determine Logik based on red text

                        if (!changeText.isEmpty()) {
                            changes.add(new ChangeInfo(tableName, changeNumber, changeText, releasestand, isFullyRed, logik));
                        }
                    }
                }
            }
        }
        return changes;
    }

    private String determineLogik(TableCell cell) {
        // Implement a heuristic approach if direct detection isn't possible
        // Placeholder implementation based on text patterns
        String cellText = cell.text().trim();
        if (cellText.contains("deleted")) { // Placeholder for detection logic
            return "Rückbau Logik";
        }
        return "Änderung Logik";
    }

    private boolean isTextFullyRed(TableCell cell) {
        // Implement logic to check if the entire text in cell is red
        // Placeholder implementation
        return cell.text().contains("red text indicator"); // Adjust as needed
    }


}
