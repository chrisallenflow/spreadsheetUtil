package com.arcanepost.service.spreadsheetutil;

import org.json.JSONObject;

public class SpreadsheetUtilResponse {

    private String status;
    private String description;
    private String data;

    public SpreadsheetUtilResponse(String status, String description, String data){
        this.status = status;
        this.description = description;
        this.data = data;
    }


    public String getStatus() {
        return status;
    }

    public String getDescription() {
        return description;
    }

    public String getData() {
        return data;
    }
}
