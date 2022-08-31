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
        HttpHeaders headers = new HttpHeaders();
        headers.setContentDispositionFormData("attachment", fileName + ".csv");
        headers.set(HttpHeaders.CONTENT_TYPE, "text/csv");

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

        return new ResponseEntity<>(new InputStreamResource(file), headers, HttpStatus.OK);
    }


}
