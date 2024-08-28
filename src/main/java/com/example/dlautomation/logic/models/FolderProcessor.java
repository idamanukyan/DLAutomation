package com.example.dlautomation.logic.models;

import com.example.dlautomation.logic.logging.GlobalLogger;
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
import java.util.logging.Level;
import java.util.logging.Logger;

public class FolderProcessor {

    private static final Logger logger = GlobalLogger.getLogger();

    public static void processFolder(String folderPath, String outputFilePath) throws IOException {

        logger.log(Level.INFO, "Processing folder: {0}", folderPath);

        File folder = new File(folderPath);
        List<File> filesToProcess = new ArrayList<>();

        int totalFileCount = collectFilesRecursively(folder, filesToProcess);

        logger.log(Level.INFO, "Number of files found to process: {0}", totalFileCount);


        if (!filesToProcess.isEmpty()) {

            logger.log(Level.INFO, "Number of files to process: {0}", filesToProcess.size());

            List<ChangeInfo> allChanges = new ArrayList<>();

            for (File file : filesToProcess) {
                String docPath = file.getAbsolutePath();
                logger.log(Level.INFO, "Processing file: {0}", docPath);
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
            logger.log(Level.INFO, "Writing changes to Excel file: {0}", outputFilePath);
            ExcelUpdater.writeChangesToExcel(allChanges, outputFilePath);
            logger.log(Level.INFO, "Process completed successfully. Output file: {0}", outputFilePath);
        } else {
            System.out.println("No documents found in the specified folder.");
            logger.log(Level.WARNING, "No documents found in the specified folder: {0}", folderPath);
        }
    }

    private static int collectFilesRecursively(File folder, List<File> filesToProcess) {
        File[] files = folder.listFiles();
        int fileCount = 0;

        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    fileCount += collectFilesRecursively(file, filesToProcess); // Recursively count files in subdirectories
                } else if (file.isFile() && (file.getName().endsWith(".doc") || file.getName().endsWith(".docx"))) {
                    filesToProcess.add(file);
                    fileCount++;
                }
            }
        }
        return fileCount;
    }


    private static List<ChangeInfo> getRedChangesWithTableName(String docPath) {
        logger.log(Level.INFO, "Getting red changes from document: {0}", docPath);
        AbstractWordReader reader;
        String fileExtension = getFileExtension(docPath);

        try {
            if (".doc".equalsIgnoreCase(fileExtension)) {
                reader = new DocWordReader(docPath);
                logger.log(Level.INFO, "Using DocWordReader for file: {0}", docPath);
            } else if (".docx".equalsIgnoreCase(fileExtension)) {
                reader = new DocxWordReader(docPath);
                logger.log(Level.INFO, "Using DocxWordReader for file: {0}", docPath);
            } else {
                logger.log(Level.SEVERE, "Unsupported file format: {0}", fileExtension);
                throw new IllegalArgumentException("Unsupported file format: " + fileExtension);
            }

            String tableName = reader.extractTableName();
            String releasestand = reader.extractReleasestand();
            return reader.getRedChanges(tableName, releasestand);
        } catch (IOException e) {
            System.err.println("Error processing file " + docPath + ": " + e.getMessage());
            logger.log(Level.SEVERE, "Error processing file " + docPath, e);
            return Collections.emptyList();
        }
    }

    private static String getFileExtension(String docPath) {
        int lastIndex = docPath.lastIndexOf('.');
        String extension = (lastIndex == -1) ? "" : docPath.substring(lastIndex);
        logger.log(Level.INFO, "File extension for {0}: {1}", new Object[]{docPath, extension});
        return extension;
    }

    public static void main(String[] args) throws IOException {
        String folderPath = "C:\\Users\\A062449\\Deutsche Leasing\\RMS-Team - Release Management\\RMS-Dokumentation\\Mappings";
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));

        Path baseFolder = Paths.get(System.getProperty("user.home"), "Downloads", "mapping_results_" + timestamp);

        File resultsFolderFile = baseFolder.toFile();
        if (!resultsFolderFile.exists()) {
            if (!resultsFolderFile.mkdirs()) {
                System.err.println("Failed to create the results folder: " + baseFolder);
                return;
            }
        }

        Path excelFilePath = baseFolder.resolve("extracted-data-" + timestamp + ".xlsx");
        Path logFilePath = baseFolder.resolve("application-" + timestamp + ".log");

        GlobalLogger.initialize(logFilePath.toString());

        logger.log(Level.INFO, "Starting folder processing with folder path: {0} and output path: {1}", new Object[]{folderPath, excelFilePath});

        processFolder(folderPath, excelFilePath.toString());

        logger.log(Level.INFO, "Processing completed. Results saved to: {0}", excelFilePath);
    }

}
