package com.arcanepost.service.spreadsheetutil;

import com.arcanepost.service.spreadsheetutil.entity.Spreadsheet;
import com.arcanepost.service.spreadsheetutil.entity.SpreadsheetRow;
import com.arcanepost.service.spreadsheetutil.entity.SpreadsheetUtilResponse;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.http.MediaType;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.context.junit4.SpringRunner;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.request.MockHttpServletRequestBuilder;
import org.springframework.test.web.servlet.request.MockMvcRequestBuilders;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;
import org.springframework.web.multipart.MultipartFile;
import java.io.*;
import java.util.ArrayList;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;

import static org.junit.Assert.*;
import static org.mockito.BDDMockito.given;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;
import static org.hamcrest.core.Is.is;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.jsonPath;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;


@RunWith(SpringRunner.class)
@WebMvcTest(SpreadsheetUtilController.class)
public class SpreadsheetUtilControllerTest {

    @Autowired
    private MockMvc mvc;

    @MockBean
    private SpreadsheetUtilController spreadsheetUtilController;

    private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetUtilControllerTest.class);

    @Test
    public void upload() {
//        ArrayList<Object> columns1 = new ArrayList<>();
//        columns1.add("test");
//        columns1.add("123");
//        columns1.add("xxx");
//        columns1.add("123");
//        columns1.add("0.32");
//        columns1.add("39.36");
//        SpreadsheetRow row1 = new SpreadsheetRow("1", columns1);
//        ArrayList<Object> columns2 = new ArrayList<>();
//        columns2.add("test2");
//        columns2.add("456");
//        columns2.add("yyy");
//        columns2.add("444");
//        columns2.add("0.44");
//        columns2.add("195.36");
//        SpreadsheetRow row2 = new SpreadsheetRow("2", columns2);
//        ArrayList<Object> columns3 = new ArrayList<>();
//        columns3.add("test3");
//        columns3.add("789");
//        columns3.add("zzz");
//        columns3.add("99");
//        columns3.add("0.99");
//        columns3.add("98.01");
//        SpreadsheetRow row3 = new SpreadsheetRow("3", columns3);
//
//        ArrayList<SpreadsheetRow> rows = new ArrayList<>();
//        rows.add(row1);
//        rows.add(row2);
//        rows.add(row3);
//        Spreadsheet sheet1 = new Spreadsheet(rows);
//        SpreadsheetUtilResponse response =
//                new SpreadsheetUtilResponse("SUCCESS", "Spreadsheet converted.", sheet1);

//        File testfile = new File("test/resources/test1.xslx");
//        LOGGER.debug("testfile.getName() = " + testfile.getName());
//        LOGGER.debug("testfile.getAbsolutePath() = " + testfile.getAbsolutePath());

        FileInputStream inputFile = null;
        try {
            MockMultipartFile file =
                    new MockMultipartFile(
                            "file",
                            "test1.xlsx",
                            null,
                            ClassLoader.getSystemResourceAsStream("test/resources/test1.xslx"));
//            given(spreadsheetUtilController.upload(file)).willReturn(response);
            MockHttpServletRequestBuilder builder =
                    MockMvcRequestBuilders
                            .multipart("/to/json")
                            .file(file);
            mvc.perform(builder).andExpect(status().isOk());

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException ee) {
            ee.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void sayHello() {
    }
}