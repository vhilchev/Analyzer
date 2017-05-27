package org.analyzer;

import java.util.HashMap;
import java.util.Map;

public class RecordEntity {
    private Map<String, String> record = new HashMap<>();
    private String columnReport;

    public void setColumnReport(String columnReport) {
        this.columnReport = columnReport;
    }

    public String getColumnReport() {
        return columnReport;
    }

    public void put(String key, String value) {
        record.put(key, value);
    }

    public String get(String key) {
        return record.get(key);
    }

    @Override
    public String toString() {
        return columnReport + ": " + record;
    }
}
