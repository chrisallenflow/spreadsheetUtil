package com.arcanepost.service.spreadsheetutil.entity;

public class SpreadsheetUtilResponse {

    private String status;
    private String description;
    private Spreadsheet data;
//    private String data;

    public SpreadsheetUtilResponse(String status, String description, Spreadsheet data){
        this.status = status;
        this.description = description;
        this.data = data;
    }

    public SpreadsheetUtilResponse(String status, String description){
        this.status = status;
        this.description = description;
    }

    public String getStatus() {
        return status;
    }

    public String getDescription() {
        return description;
    }

    public Spreadsheet getData() {
        return data;
    }
}
