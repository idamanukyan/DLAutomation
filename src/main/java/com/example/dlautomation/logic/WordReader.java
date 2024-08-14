/*
package com.example.dlautomation.logic;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class WordReader {

    public static List<ChangeInfo> getRedChangesWithTableName(String docPath) throws IOException {
        List<ChangeInfo> changes = new ArrayList<>();
        String fileExtension = getFileExtension(docPath);

        String tableName;
        String releasestand;

        if (".doc".equalsIgnoreCase(fileExtension)) {
            tableName = extractTableNameFromDoc(docPath);
            releasestand = extractReleasestandFromDoc(docPath);
            changes.addAll(getRedChangesFromDoc(docPath, tableName, releasestand));
        } else if (".docx".equalsIgnoreCase(fileExtension)) {
            tableName = extractTableNameFromDocx(docPath);
            releasestand = extractReleasestandFromDocx(docPath);
            changes.addAll(getRedChangesFromDocx(docPath, tableName, releasestand));
        } else {
            throw new IllegalArgumentException("Unsupported file format: " + fileExtension);
        }

        return changes;
    }


    private static String extractTableNameFromDoc(String docPath) throws IOException {
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

    private static String extractTableNameFromDocx(String docPath) throws IOException {
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

    private static String extractReleasestandFromDoc(String docPath) throws IOException {
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

    private static String extractReleasestandFromDocx(String docPath) throws IOException {
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



    private static List<ChangeInfo> getRedChangesFromDoc(String docPath, String tableName, String releasestand) throws IOException {
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

                        // Add change info if valid
                        if (!changeText.isEmpty() && isRed(changeText)) {
                            changes.add(new ChangeInfo(tableName, changeNumber, changeText, releasestand));
                        }
                    }
                }
            }
        }

        return changes;
    }

    private static List<ChangeInfo> getRedChangesFromDocx(String docPath, String tableName, String releasestand) throws IOException {
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

                        // Add change info if valid
                        if (!changeText.isEmpty()) {
                            changes.add(new ChangeInfo(tableName, changeNumber, changeText, releasestand));
                        }
                    }
                }
            }
        }

        return changes;
    }

    private static String getRedTextFromCell(XWPFTableCell cell) {
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

    private static boolean isRed(String text) {
        // Implement logic to determine if the text is red. This might be required if text color information is embedded in the text.
        return text.contains("some red text identifier"); // This is a placeholder
    }

    private static String getFileExtension(String docPath) {
        int lastIndex = docPath.lastIndexOf('.');
        return (lastIndex == -1) ? "" : docPath.substring(lastIndex);
    }

    public static void main(String[] args) throws IOException {
        String docPath = "C:\\Users\\Admin\\Downloads\\MOD_RMSOUT_ABACUS.MAP_T_ABACUS_RT930_COLL (1).docx";

        // Extract the releasestand value
        String releasestand = extractReleasestandFromDoc(docPath);

        // Generate a timestamp for the file name
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String excelPath = "C:\\Users\\Admin\\Downloads\\RedChanges_" + timestamp + ".xlsx";

        List<ChangeInfo> changes = getRedChangesWithTableName(docPath);

        // Print the extracted red changes
        System.out.println("Red Changes:");
        for (ChangeInfo change : changes) {
            System.out.println("Table Name: " + change.getTableName() + " | Change Number: " + change.getChangeNumber() + " | Change: " + change.getChange());
        }

        // Write the changes to an Excel file with releasestand
        ExcelUpdater.writeChangesToExcel(changes, excelPath, releasestand);
    }

}
*/
