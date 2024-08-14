package com.example.dlautomation.logic.inheritance;

import com.example.dlautomation.logic.ChangeInfo;

import java.io.IOException;
import java.util.List;

public abstract class AbstractWordReader {

    protected String docPath;

    public AbstractWordReader(String docPath) {
        this.docPath = docPath;
    }

    public abstract String extractTableName() throws IOException;

    public abstract String extractReleasestand() throws IOException;

    public abstract List<ChangeInfo> getRedChanges(String tableName, String releasestand) throws IOException;

    protected String getFileExtension() {
        int lastIndex = docPath.lastIndexOf('.');
        return (lastIndex == -1) ? "" : docPath.substring(lastIndex);
    }
}
