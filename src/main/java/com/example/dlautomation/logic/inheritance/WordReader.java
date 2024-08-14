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

    public static void main(String[] args) throws IOException {
        String docPath = "C:\\Users\\Admin\\Downloads\\MOD_RMSIN_DEAL.MAP_T_ORBI_DEALS_AFTERMATH1.docx" ;

        List<ChangeInfo> changes = getRedChangesWithTableName(docPath);

        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String excelPath = "C:\\Users\\Admin\\Downloads\\RedChanges_" + timestamp + ".xlsx";

        // Print the extracted red changes
        System.out.println("Red Changes:");
        for (ChangeInfo change : changes) {
            System.out.println("Table Name: " + change.getTableName() + " | Change Number: " + change.getChangeNumber() + " | Change: " + change.getChange());
        }

        // Write the changes to an Excel file
        ExcelUpdater.writeChangesToExcel(changes, excelPath, changes.get(0).getReleasestand());
    }
}
