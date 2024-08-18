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
        }
        return "Unknown Table Name";
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

                        boolean isFullyRed = getRedTextFromCell(changeCell).equals(changeCell.getText());
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

    private String determineLogik(XWPFTableCell cell) {
        StringBuilder cellText = new StringBuilder();
        boolean isCrossedOut = false;

        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                if (run.getColor() != null && "FF0000".equalsIgnoreCase(run.getColor())) {
                    cellText.append(run.getText(0)).append(" ");
                    if (run.isStrikeThrough()) {
                        isCrossedOut = true;
                    }
                }
            }
        }

        if (isCrossedOut) {
            return "Rückbau Logik";
        }
        return "Neuer Variable";
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
}
