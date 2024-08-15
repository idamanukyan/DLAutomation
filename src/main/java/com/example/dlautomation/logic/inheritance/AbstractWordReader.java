package com.example.dlautomation.logic.inheritance;

import com.example.dlautomation.logic.ChangeInfo;

import java.io.IOException;
import java.util.List;

public abstract class AbstractWordReader {

    protected String docPath;
    protected String module;
    protected String mapping;

    public AbstractWordReader(String docPath) {
        this.docPath = docPath;
        extractModuleAndMapping();
    }

    public abstract String extractTableName() throws IOException;

    public abstract String extractReleasestand() throws IOException;

    public abstract List<ChangeInfo> getRedChanges(String tableName, String releasestand) throws IOException;

    protected String getFileExtension() {
        int lastIndex = docPath.lastIndexOf('.');
        return (lastIndex == -1) ? "" : docPath.substring(lastIndex);
    }

    private void extractModuleAndMapping() {
        // Extract module and mapping from the filename
        String fileName = docPath.substring(docPath.lastIndexOf("\\") + 1, docPath.lastIndexOf('.'));
        String[] parts = fileName.split("\\.");

        if (parts.length >= 2) {
            String modulePart = parts[0];
            String mappingPart = parts[1];

            if (modulePart.startsWith("MOD_")) {
                module = modulePart.substring(4); // Extract module
            }
            if (mappingPart.startsWith("MAP_")) {
                mapping = mappingPart.substring(4); // Extract mapping
            }
        }
    }

    public String getModule() {
        return module;
    }

    public String getMapping() {
        return mapping;
    }
}
