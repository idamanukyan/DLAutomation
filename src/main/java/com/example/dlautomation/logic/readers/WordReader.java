package com.example.dlautomation.logic.readers;

import com.example.dlautomation.logic.models.ChangeInfo;

import java.io.IOException;
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
}
