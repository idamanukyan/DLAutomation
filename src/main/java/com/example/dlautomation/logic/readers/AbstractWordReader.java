package com.example.dlautomation.logic.readers;

import com.example.dlautomation.logic.logging.GlobalLogger;
import com.example.dlautomation.logic.models.ChangeInfo;

import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

public abstract class AbstractWordReader {

    private static final Logger logger = GlobalLogger.getLogger();

    protected String docPath;
    protected String module;
    protected String mappingName;

    public AbstractWordReader(String docPath) {
        this.docPath = docPath;
        logger.log(Level.INFO, "Initialized AbstractWordReader with document path: {0}", docPath);
        extractModuleAndMapping();
    }

    public abstract String extractTableName() throws IOException;

    public abstract String extractReleasestand() throws IOException;

    public abstract List<ChangeInfo> getRedChanges(String tableName, String releasestand) throws IOException;

    private void extractModuleAndMapping() {
        logger.log(Level.INFO, "Extracting module and mapping from document path: {0}", docPath);
        String fileName = docPath.substring(docPath.lastIndexOf("\\") + 1, docPath.lastIndexOf('.'));
        String[] parts = fileName.split("\\.");

        if (parts.length >= 2) {
            String modulePart = parts[0];
            String mappingPart = parts[1];

            if (modulePart.startsWith("MOD_")) {
                module = modulePart.substring(4);
                logger.log(Level.INFO, "Extracted module: {0}", module);
            }
            if (mappingPart.startsWith("MAP_")) {
                mappingName = mappingPart.substring(4);
                logger.log(Level.INFO, "Extracted mapping name: {0}", mappingName);
            } else {
                logger.log(Level.WARNING, "Mapping name not found in file name: {0}", fileName);
            }
        } else {
            logger.log(Level.WARNING, "Unexpected file name format: {0}", fileName);
        }
    }

    public String getMappingName() {
        logger.log(Level.INFO, "Returning mapping name: {0}", mappingName);
        return mappingName;
    }
}
