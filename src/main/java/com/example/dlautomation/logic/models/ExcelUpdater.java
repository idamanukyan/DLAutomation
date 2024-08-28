package com.example.dlautomation.logic.models;

import com.example.dlautomation.logic.logging.GlobalLogger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;


public class ExcelUpdater {

    private static final Logger logger = GlobalLogger.getLogger();

    public static void writeChangesToExcel(List<ChangeInfo> changes, String filePath) throws IOException {

        logger.log(Level.INFO, "Starting to write changes to Excel. File path: {0}", filePath);

        try (Workbook workbook = new XSSFWorkbook()) {

            logger.log(Level.INFO, "Workbook created successfully.");

            Sheet datenmodellanderungenSheet = workbook.createSheet("datenmodellanderungen");
            Sheet logikanderungenSheet = workbook.createSheet("logikanderungen");

            logger.log(Level.INFO, "Sheets created: datenmodellanderungen and logikanderungen.");

            createHeaderRow(datenmodellanderungenSheet);
            createHeaderRow(logikanderungenSheet);

            logger.log(Level.INFO, "Header rows created in both sheets.");

            int datenmodellanderungenRowNum = 1;
            int logikanderungenRowNum = 1;

            for (ChangeInfo change : changes) {
                Row row;
                if (change.isFullyRed()) {
                    row = datenmodellanderungenSheet.createRow(datenmodellanderungenRowNum++);
                    logger.log(Level.INFO, "Writing change to datenmodellanderungen sheet: {0}", change);
                } else {
                    row = logikanderungenSheet.createRow(logikanderungenRowNum++);
                    logger.log(Level.INFO, "Writing change to logikanderungen sheet: {0}", change);
                }

                createDataRow(row, change);
            }

            logger.log(Level.INFO, "All changes have been written to the sheets.");

            workbook.setSheetName(0, "Datenmodelländerungen");
            workbook.setSheetName(1, "Logikänderungen");

            logger.log(Level.INFO, "Sheet names updated: Datenmodelländerungen and Logikänderungen.");

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
                logger.log(Level.INFO, "Workbook written to file successfully. File path: {0}", filePath);
            }
        } catch (IOException e) {
            logger.log(Level.SEVERE, "IOException occurred while writing to Excel file.", e);
            throw e;
        }
        logger.log(Level.INFO, "Excel writing process completed.");
    }

    private static void createHeaderRow(Sheet sheet) {

        logger.log(Level.INFO, "Creating header row for sheet: {0}", sheet.getSheetName());

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Tabellenname");
        headerRow.createCell(1).setCellValue("Feldname");
        headerRow.createCell(2).setCellValue("Änderung");
        headerRow.createCell(3).setCellValue("Releasestand");
        headerRow.createCell(4).setCellValue("Logik");
        headerRow.createCell(5).setCellValue("Mappingname");
        headerRow.createCell(6).setCellValue("Ganze Reihe");

        logger.log(Level.INFO, "Header row created successfully.");

    }

    private static void createDataRow(Row row, ChangeInfo change) {

        logger.log(Level.INFO, "Creating data row for change: {0}", change);

        row.createCell(0).setCellValue(change.getTableName());
        row.createCell(1).setCellValue(change.getChangeNumber());
        row.createCell(2).setCellValue(change.getChange());
        row.createCell(3).setCellValue(change.getReleasestand());
        row.createCell(4).setCellValue(change.getLogik());
        row.createCell(5).setCellValue(change.getMappingName());
        row.createCell(6).setCellValue(change.getWholeString());

        logger.log(Level.INFO, "Data row created successfully for change: {0}", change);
    }
}
