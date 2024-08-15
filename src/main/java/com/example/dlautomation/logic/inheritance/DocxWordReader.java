package com.example.dlautomation.logic.inheritance;

import com.example.dlautomation.logic.ChangeInfo;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class DocxWordReader extends AbstractWordReader {

    private String module;
    private String mapping;

    public DocxWordReader(String docPath) {
        super(docPath);
        extractModuleAndMapping();
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
                            changes.add(new ChangeInfo(tableName, changeNumber, changeText, releasestand, module, mapping, isFullyRed, logik));
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
        return "Änderung Logik";
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

    private void extractModuleAndMapping() {
        String fileName = docPath.substring(docPath.lastIndexOf("\\") + 1, docPath.lastIndexOf('.'));
        String[] parts = fileName.split("\\.");
        if (parts.length >= 3) {
            String modulePart = parts[0].toUpperCase();
            String mappingPart = parts[1].toUpperCase();
            this.module = modulePart.split("_")[1]; // Assuming the format MOD_MODULE
            this.mapping = mappingPart.split("_")[1]; // Assuming the format MAP_MAPPING
        } else {
            this.module = "Unknown Module";
            this.mapping = "Unknown Mapping";
        }
    }
}
