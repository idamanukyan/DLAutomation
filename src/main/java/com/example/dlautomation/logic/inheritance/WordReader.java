package com.example.dlautomation.logic.inheritance;

import com.example.dlautomation.logic.ChangeInfo;
import com.example.dlautomation.logic.ExcelUpdater;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class WordReader {

    public static List<ChangeInfo> getRedChangesWithTableName(String docPath) throws IOException {
        AbstractWordReader reader;
        String fileExtension = getFileExtension(docPath);

        if (".doc".equalsIgnoreCase(fileExtension)) {
            reader = new DocWordReader(docPath);
        } else if (".docx".equalsIgnoreCase(fileExtension)) {
            reader = new DocxWordReader(docPath);
        } else {
            throw new IllegalArgumentException("Unsupported file format: " + fileExtension);
        }

        String tableName = reader.extractTableName();
        String releasestand = reader.extractReleasestand();
        return reader.getRedChanges(tableName, releasestand);
    }

    private static String getFileExtension(String docPath) {
        int lastIndex = docPath.lastIndexOf('.');
        return (lastIndex == -1) ? "" : docPath.substring(lastIndex);
    }

    private static String extractModule(String fileName) {
        String module = "";
        int modIndex = fileName.toUpperCase().indexOf("MOD_");
        int mapIndex = fileName.toUpperCase().indexOf(".MAP");

        if (modIndex != -1 && mapIndex != -1) {
            int startIndex = modIndex + 4;
            module = fileName.substring(startIndex, mapIndex).trim();
        }

        return module;
    }

    private static String extractMapping(String fileName) {
        String mapping = "";
        int mapIndex = fileName.toUpperCase().indexOf(".MAP");
        int extensionIndex = fileName.lastIndexOf('.');

        if (mapIndex != -1 && extensionIndex != -1 && extensionIndex > mapIndex) {
            int startIndex = mapIndex + 4;
            mapping = fileName.substring(startIndex, extensionIndex).trim();
        }

        return mapping;
    }

    public static void main(String[] args) throws IOException {
        String docPath = "C:\\Users\\Admin\\Downloads\\MOD_RMSIN_DEAL.MAP_T_ORBI_DEALS_AFTERMATH1.docx";
        List<ChangeInfo> changes = getRedChangesWithTableName(docPath);

        // Extract module and mapping from the file name
        String fileName = docPath.substring(docPath.lastIndexOf("\\") + 1);
        String module = extractModule(fileName);
        String mapping = extractMapping(fileName);

        // Use module and mapping to name the Excel file
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String excelPath = String.format("C:\\Users\\Admin\\Downloads\\%s_%s_Changes_%s.xlsx", module, mapping, timestamp);

        System.out.println("Red Changes:");
        for (ChangeInfo change : changes) {
            System.out.println("Table Name: " + change.getTableName() + " | Change Number: " + change.getChangeNumber() + " | Change: " + change.getChange());
        }

        ExcelUpdater.writeChangesToExcel(changes, excelPath, module, mapping);
    }
}
