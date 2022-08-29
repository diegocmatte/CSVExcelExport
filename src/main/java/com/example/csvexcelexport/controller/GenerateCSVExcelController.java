package com.example.csvexcelexport.controller;


import com.example.csvexcelexport.pojo.ClientDataObjectRequest;
import com.example.csvexcelexport.service.GenerateCSVExcelService;
import io.swagger.annotations.Api;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.client.RestTemplate;

import java.io.*;
import java.text.SimpleDateFormat;


@RestController
@RequestMapping("/generate")
@Api(value = "Generate files controller")
public class GenerateCSVExcelController {

    private static final Logger log = LogManager.getLogger(GenerateCSVExcelController.class);

    @Autowired
    private GenerateCSVExcelService generateCSVExcelService;


    @PostMapping(value = "/exportfile/csv", produces = {"text/csv"}, consumes = {MediaType.APPLICATION_JSON_VALUE})
    public ResponseEntity<InputStreamResource> exportCSV(@RequestBody ClientDataObjectRequest clientDataObjectRequest) throws IOException {

        log.info("Creating file name.");
        String dateFile = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String timeFile = new SimpleDateFormat("HH-mm").format(new java.util.Date());
        String fileName = clientDataObjectRequest.getMetricName().replaceAll("[ /]","_") + "_" + dateFile + "T" + timeFile;
        log.info("file " + fileName + " created.");

        log.info("Calling the service class.");
        InputStreamResource fileInputStream = generateCSVExcelService.generateCSV(clientDataObjectRequest);

        if(!fileInputStream.exists()){
            log.error("Error while creating file.");
            return new ResponseEntity<>(null, null, HttpStatus.NO_CONTENT);
        }

        log.info("Creating HttpHeaders");
        System.out.println(jsonMock());
        HttpHeaders headers = new HttpHeaders();
        headers.setContentDispositionFormData("attachment", fileName + ".csv");
        headers.set(HttpHeaders.CONTENT_TYPE, "text/csv");
        //headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName + ".csv");
        //headers.set(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_OCTET_STREAM_VALUE);

        return new ResponseEntity<>(
                fileInputStream,
                headers,
                HttpStatus.OK
        );
    }

    @PostMapping(value = "/exportfile/excel", produces = {MediaType.APPLICATION_OCTET_STREAM_VALUE}, consumes = {MediaType.APPLICATION_JSON_VALUE})
    public ResponseEntity<InputStreamResource> exportExcel(@RequestBody ClientDataObjectRequest clientDataObjectRequest) {

        ByteArrayInputStream file = generateCSVExcelService.generateFileDetailsWithChart(clientDataObjectRequest);

        if(file.available() < 1){
            log.error("Error while creating file.");
            return new ResponseEntity<>(null, null, HttpStatus.NO_CONTENT);
        }

        String dateFile = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String timeFile = new SimpleDateFormat("HH-mm").format(new java.util.Date());
        String fileName = clientDataObjectRequest.getMetricName().replaceAll("[ /]","_") + "_" + dateFile + "T" + timeFile;

        HttpHeaders headers = new HttpHeaders();
        headers.setContentDispositionFormData("attachment", fileName + ".xlsx");
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        //headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName + ".xlsx");
        //headers.set(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_OCTET_STREAM_VALUE);

        return new ResponseEntity<>(new InputStreamResource(file), headers, HttpStatus.OK);
    }

    private String jsonMock(){
        return "{\n" +
                "    \"metricName\":\"Turnover/Rate abc\",\n" +
                "    \"dataFormatCodeValue\": \"currency\",\n" +
                "    \"clientDataRequest\":[\n" +
                "       {\n" +
                "          \"clientName\":\"client 1\",\n" +
                "          \"value\":\"8\"\n" +
                "       },\n" +
                "       {\n" +
                "          \"clientName\":\"client 2\",\n" +
                "          \"value\":\"7\"\n" +
                "       },\n" +
                "       {\n" +
                "          \"clientName\":\"client 3\",\n" +
                "          \"value\":\"6\"\n" +
                "       },\n" +
                "       {\n" +
                "          \"clientName\":\"client 4\",\n" +
                "          \"value\":\"5\"\n" +
                "       },\n" +
                "       {\n" +
                "          \"clientName\":\"client 5555555\",\n" +
                "          \"value\":\"4\"\n" +
                "       },\n" +
                "       {\n" +
                "          \"clientName\":\"client 6\",\n" +
                "          \"value\":\"3\"\n" +
                "       },\n" +
                "       {\n" +
                "          \"clientName\":\"client 7\",\n" +
                "          \"value\":\"2\"\n" +
                "       },\n" +
                "       {\n" +
                "          \"clientName\":\"client 8\",\n" +
                "          \"value\":\"1\"\n" +
                "       }\n" +
                "    ]\n" +
                "}";
    }

}
