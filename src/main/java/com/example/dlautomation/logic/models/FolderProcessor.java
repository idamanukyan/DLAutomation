package com.example.dlautomation.logic.models;

import com.example.dlautomation.logic.readers.AbstractWordReader;
import com.example.dlautomation.logic.readers.DocWordReader;
import com.example.dlautomation.logic.readers.DocxWordReader;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class FolderProcessor {

    public static void processFolder(String folderPath, String outputFilePath) throws IOException {
        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.endsWith(".doc") || name.endsWith(".docx"));

        if (files != null) {
            List<ChangeInfo> allChanges = new ArrayList<>();

            for (File file : files) {
                String docPath = file.getAbsolutePath();
                List<ChangeInfo> changes = getRedChangesWithTableName(docPath);

                // Add changes to the list
                allChanges.addAll(changes);
            }

            // Use dummy values for module and mapping since it's for all documents
            String dummyModule = "";
            String dummyMapping = "";

            // Write all changes to a single Excel file
            ExcelUpdater.writeChangesToExcel(allChanges, outputFilePath, dummyModule, dummyMapping);
        } else {
            System.out.println("No documents found in the specified folder.");
        }
    }

    private static List<ChangeInfo> getRedChangesWithTableName(String docPath) throws IOException {
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
        String folderPath = "C:\\Users\\Admin\\Desktop\\documents";
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String outputFilePath = String.format("C:\\Users\\Admin\\Downloads\\All_Changes_%s.xlsx", timestamp);

        processFolder(folderPath, outputFilePath);
    }
}
