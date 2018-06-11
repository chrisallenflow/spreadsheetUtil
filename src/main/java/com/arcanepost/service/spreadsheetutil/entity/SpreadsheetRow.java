package com.arcanepost.service.spreadsheetutil.entity;

import java.util.ArrayList;
import java.util.List;

public class SpreadsheetRow {

    private String rowNumber;
    private List<Object> columns;

    public SpreadsheetRow(String rowNumber, List<Object> columns) {

        this.rowNumber = rowNumber;
        this.columns = columns;
    }

    public List<Object> getColumns() {
        return columns;
    }

    public String getRowNumber() {
        return rowNumber;
    }
}
