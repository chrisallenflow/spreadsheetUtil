package com.arcanepost.service.spreadsheetutil.entity;

import java.util.ArrayList;

public class Spreadsheet {

    private ArrayList<SpreadsheetRow> rows;

    public Spreadsheet(ArrayList<SpreadsheetRow> rows) {
        this.rows = rows;
    }

    public ArrayList<SpreadsheetRow> getRows() {
        return rows;
    }
}
