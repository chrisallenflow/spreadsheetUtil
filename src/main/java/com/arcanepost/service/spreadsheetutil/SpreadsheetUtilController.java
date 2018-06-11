package com.arcanepost.service.spreadsheetutil;

import com.arcanepost.service.spreadsheetutil.entity.Spreadsheet;
import com.arcanepost.service.spreadsheetutil.entity.SpreadsheetRow;
import com.arcanepost.service.spreadsheetutil.entity.SpreadsheetUtilResponse;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.Iterator;

@RestController
public class SpreadsheetUtilController {

    private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetUtilController.class);

    @RequestMapping(value = "/to/json", method = RequestMethod.POST)
    public SpreadsheetUtilResponse upload(MultipartFile file) {
//        public SpreadsheetUtilResponse upload(@RequestParam("file") MultipartFile file) {

        LOGGER.info("in spreadsheetToJson...");
        XSSFWorkbook workbook = null;

        if (file == null) {
            return new SpreadsheetUtilResponse(
                    "ERROR",
                    "Make sure that the spreadsheet you post is to a parameter with name 'file'." );
        }

        try {
            LOGGER.debug("file.getOriginalFilename() = " + file.getOriginalFilename());

            if (file.getOriginalFilename().endsWith("xls")) {
                LOGGER.info("about to convert xls to xlsx");
                workbook = new XSSFWorkbook(ApachePOIUtil.convertXlsToXlsx(file.getInputStream()));
            } else if (file.getOriginalFilename().endsWith("csv")) {
                LOGGER.info("about to convert csv to xlsx");
                workbook = new XSSFWorkbook(ApachePOIUtil.convertCsvToXlsx(file.getInputStream()));
            } else if (file.getOriginalFilename().endsWith("xlsx")) {
                workbook = new XSSFWorkbook(file.getInputStream());
            } else {
                return new SpreadsheetUtilResponse(
                        "ERROR",
                        "File extension must be xls, xlsx, or csv.");
            }
            JSONObject json = ApachePOIUtil.convertXlsxToJson(workbook);
            Spreadsheet sheet = convertJsonToEntity(json);

            return new SpreadsheetUtilResponse("SUCCESS", "Spreadsheet converted.", sheet);

        } catch (Exception ex) {
            return new SpreadsheetUtilResponse(
                    "ERROR",
                    ex.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (Exception e) {
                LOGGER.error("workbook could not be closed in finally clause....ignoring.");
            }
        }
    }

    @RequestMapping(value = "/sayHello", method = RequestMethod.GET)
    public SpreadsheetUtilResponse sayHello(@RequestParam("username") String username) {

        LOGGER.debug("*** about to say hello ***");
        try {
            return new SpreadsheetUtilResponse(
                    "SUCCESS",
                    "Hello, " + username + "!  Welcome!!");

        } catch (Exception e) {
            LOGGER.error("*** something went wrong....");
            return new SpreadsheetUtilResponse(
                    "ERROR",
                    "Could not say hello for some reason" + e.getMessage());
        }

    }

    private Spreadsheet convertJsonToEntity(JSONObject json){

        ArrayList<SpreadsheetRow> rows = new ArrayList<>();

        Iterator<String> rowIter = json.keys();
        while (rowIter.hasNext()){
            String key = rowIter.next();
            JSONArray array = json.getJSONArray(key);
            SpreadsheetRow row = new SpreadsheetRow(key, array.toList());
            rows.add(row);
        }

        return new Spreadsheet(rows);

    }
}
