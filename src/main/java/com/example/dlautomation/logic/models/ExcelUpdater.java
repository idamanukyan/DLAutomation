package com.example.dlautomation.logic.models;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelUpdater {

    public static void writeChangesToExcel(List<ChangeInfo> changes, String filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet datenmodellanderungenSheet = workbook.createSheet("datenmodellanderungen");
            Sheet logikanderungenSheet = workbook.createSheet("logikanderungen");

            createHeaderRow(datenmodellanderungenSheet);
            createHeaderRow(logikanderungenSheet);

            int datenmodellanderungenRowNum = 1;
            int logikanderungenRowNum = 1;

            for (ChangeInfo change : changes) {
                Row row;
                if (change.isFullyRed()) {
                    row = datenmodellanderungenSheet.createRow(datenmodellanderungenRowNum++);
                } else {
                    row = logikanderungenSheet.createRow(logikanderungenRowNum++);
                }

                createDataRow(row, change);
            }

            workbook.setSheetName(0, "Datenmodelländerungen");
            workbook.setSheetName(1, "Logikänderungen");

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
        }
    }

    private static void createHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Tabellenname");
        headerRow.createCell(1).setCellValue("Feldname");
        headerRow.createCell(2).setCellValue("Änderung");
        headerRow.createCell(3).setCellValue("Releasestand");
        headerRow.createCell(4).setCellValue("Logik");
        headerRow.createCell(5).setCellValue("Mappingname");
        headerRow.createCell(6).setCellValue("Whole Text");
    }

    private static void createDataRow(Row row, ChangeInfo change) {
        row.createCell(0).setCellValue(change.getTableName());
        row.createCell(1).setCellValue(change.getChangeNumber());
        row.createCell(2).setCellValue(change.getChange());
        row.createCell(3).setCellValue(change.getReleasestand());
        row.createCell(4).setCellValue(change.getLogik());
        row.createCell(5).setCellValue(change.getMappingName());
        row.createCell(6).setCellValue(change.getWholeString());
    }
}
