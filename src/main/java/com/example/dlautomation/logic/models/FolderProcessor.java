package com.example.dlautomation.logic.models;

import com.example.dlautomation.logic.readers.AbstractWordReader;
import com.example.dlautomation.logic.readers.DocWordReader;
import com.example.dlautomation.logic.readers.DocxWordReader;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class FolderProcessor {

    public static void processFolder(String folderPath, String outputFilePath) throws IOException {
        File folder = new File(folderPath);
        List<File> filesToProcess = new ArrayList<>();

        collectFilesRecursively(folder, filesToProcess);

        if (!filesToProcess.isEmpty()) {
            List<ChangeInfo> allChanges = new ArrayList<>();

            for (File file : filesToProcess) {
                String docPath = file.getAbsolutePath();
                String mappingName = file.getName().substring(0, file.getName().lastIndexOf('.'));

                List<ChangeInfo> notFilteredChanges = getRedChangesWithTableName(docPath);

                List<ChangeInfo> changes = notFilteredChanges.stream()
                        .filter(change -> !"Join-Bedingungen".equalsIgnoreCase(change.getChangeNumber()))
                        .toList();

                for (ChangeInfo change : changes) {
                    allChanges.add(new ChangeInfo(
                            change.getTableName(),
                            change.getChangeNumber(),
                            change.getChange(),
                            change.getReleasestand(),
                            mappingName,
                            change.isFullyRed(),
                            change.getLogik(),
                            change.getWholeString()
                    ));
                }
            }

            ExcelUpdater.writeChangesToExcel(allChanges, outputFilePath);
        } else {
            System.out.println("No documents found in the specified folder.");
        }
    }

    private static void collectFilesRecursively(File folder, List<File> filesToProcess) {
        File[] files = folder.listFiles();

        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    collectFilesRecursively(file, filesToProcess);
                } else if (file.isFile() && (file.getName().endsWith(".doc") || file.getName().endsWith(".docx"))) {
                    filesToProcess.add(file);
                }
            }
        }
    }

    private static List<ChangeInfo> getRedChangesWithTableName(String docPath) {
        AbstractWordReader reader;
        String fileExtension = getFileExtension(docPath);

        try {
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
        } catch (IOException e) {
            System.err.println("Error processing file " + docPath + ": " + e.getMessage());
            return Collections.emptyList();
        }
    }

    private static String getFileExtension(String docPath) {
        int lastIndex = docPath.lastIndexOf('.');
        return (lastIndex == -1) ? "" : docPath.substring(lastIndex);
    }

    public static void main(String[] args) throws IOException {
        String folderPath = "C:\\Users\\A062449\\Deutsche Leasing\\RMS-Team - Release Management\\RMS-Dokumentation\\Mappings";
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));

        Path downloadsPath = Paths.get(System.getProperty("user.home"), "Downloads", "extracted-data-" + timestamp + ".xlsx");

        processFolder(folderPath, downloadsPath.toString());
    }
}
