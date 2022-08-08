package com.example.csvexcelexport.controller;


import com.example.csvexcelexport.pojo.ClientDataObjectRequest;
import com.example.csvexcelexport.service.GenerateCSVExcelService;
import io.swagger.annotations.Api;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.text.SimpleDateFormat;

@RestController
@RequestMapping("/generate")
@Api(value = "Generate files controller")
public class GenerateCSVExcelController {

    @Autowired
    private GenerateCSVExcelService generateCSVExcelService;

    @PostMapping(value = "/exportfile/csv")
    public ResponseEntity<InputStreamResource> exportCSV(@RequestBody ClientDataObjectRequest clientDataObjectRequest) {

        String dateFile = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String timeFile = new SimpleDateFormat("HH-mm").format(new java.util.Date());
        String fileName = clientDataObjectRequest.getMetricName().replace(" ","_") + "_" + dateFile + "T" + timeFile;

        InputStreamResource fileInputStream = generateCSVExcelService.generateCSV(clientDataObjectRequest);

        HttpHeaders headers = new HttpHeaders();
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName + ".csv");
        headers.set(HttpHeaders.CONTENT_TYPE, "text/csv");

        return new ResponseEntity<>(
                fileInputStream,
                headers,
                HttpStatus.OK
        );
    }

    @PostMapping(value = "/exportfile/excel")
    public ResponseEntity<InputStreamResource> exportExcel(@RequestBody ClientDataObjectRequest clientDataObjectRequest) {

        ByteArrayInputStream file = generateCSVExcelService.generateFileDetailsWithChart(clientDataObjectRequest);

        String dateFile = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String timeFile = new SimpleDateFormat("HH-mm").format(new java.util.Date());
        String fileName = clientDataObjectRequest.getMetricName().replace(" ","_") + "_" + dateFile + "T" + timeFile;

        HttpHeaders headers = new HttpHeaders();
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName + ".xlsx");
        headers.set(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_JSON_VALUE);

        return new ResponseEntity<>(new InputStreamResource(file), headers, HttpStatus.OK);
    }

}
