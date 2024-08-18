package com.example.dlautomation.logic.models;

public class ChangeInfo {
    private String tableName;
    private String changeNumber;
    private String change;
    private String releasestand;
    private String mappingName;
    private boolean isFullyRed;
    private String logik;

    public ChangeInfo(String tableName, String changeNumber, String change, String releasestand, String mappingName, boolean isFullyRed, String logik) {
        this.tableName = tableName;
        this.changeNumber = changeNumber;
        this.change = change;
        this.releasestand = releasestand;
        this.mappingName = mappingName;
        this.isFullyRed = isFullyRed;
        this.logik = logik;
    }

    public String getTableName() {
        return tableName;
    }

    public String getChangeNumber() {
        return changeNumber;
    }

    public String getChange() {
        return change;
    }

    public String getReleasestand() {
        return releasestand;
    }

    public String getMappingName() {
        return mappingName;
    }

    public boolean isFullyRed() {
        return isFullyRed;
    }

    public String getLogik() {
        return logik;
    }
}
