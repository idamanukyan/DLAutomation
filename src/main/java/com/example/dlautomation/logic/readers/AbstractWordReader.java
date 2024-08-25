package com.example.dlautomation.logic.readers;

import com.example.dlautomation.logic.models.ChangeInfo;

import java.io.IOException;
import java.util.List;

public abstract class AbstractWordReader {

    protected String docPath;
    protected String module;
    protected String mappingName;

    public AbstractWordReader(String docPath) {
        this.docPath = docPath;
        extractModuleAndMapping();
    }

    public abstract String extractTableName() throws IOException;

    public abstract String extractReleasestand() throws IOException;

    public abstract List<ChangeInfo> getRedChanges(String tableName, String releasestand) throws IOException;

    private void extractModuleAndMapping() {
        String fileName = docPath.substring(docPath.lastIndexOf("\\") + 1, docPath.lastIndexOf('.'));
        String[] parts = fileName.split("\\.");

        if (parts.length >= 2) {
            String modulePart = parts[0];
            String mappingPart = parts[1];

            if (modulePart.startsWith("MOD_")) {
                module = modulePart.substring(4);
            }
            if (mappingPart.startsWith("MAP_")) {
                mappingName = mappingPart.substring(4);
            }
        }
    }

    public String getMappingName() {
        return mappingName;
    }
}
