package com.example.csvexcelexport.controller;


import com.example.csvexcelexport.pojo.ClientDataObjectRequest;
import com.example.csvexcelexport.pojo.CompanyChartData;
import com.example.csvexcelexport.service.GenerateCSVExcelService;
import io.swagger.annotations.Api;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.QuoteMode;
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
import org.supercsv.io.CsvBeanWriter;
import org.supercsv.io.ICsvBeanWriter;
import org.supercsv.prefs.CsvPreference;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/generate")
@Api(value = "Generate files controller")
public class GenerateCSVExcelController {

    @Autowired
    private GenerateCSVExcelService generateCSVExcelService;

    @PostMapping("/users/export")
    public void exportToCSV(@RequestBody ClientDataObjectRequest clientDataObjectRequest, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");

        String dateFile = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String timeFile = new SimpleDateFormat("HH-mm").format(new java.util.Date());
        String fileName = clientDataObjectRequest.getMetricName().replace(" ","_") + "_" + dateFile + "T" + timeFile;

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=" + fileName + ".csv";
        response.setHeader(headerKey, headerValue);

        List<CompanyChartData> listUsers = clientDataObjectRequest.getCompanyChartData();

        ICsvBeanWriter csvWriter = new CsvBeanWriter(response.getWriter(), CsvPreference.EXCEL_PREFERENCE);
        String[] csvHeader = {"Client Name","Value"};
        String[] nameMapping = {"clientName","value"};

        csvWriter.writeHeader(csvHeader);

        for (CompanyChartData user : listUsers) {
            csvWriter.write(user, nameMapping);
        }

        csvWriter.close();

    }

    @PostMapping(value = "/v1/csv")
    public ResponseEntity<InputStreamResource> exportCSVV2(@RequestBody ClientDataObjectRequest clientDataObjectRequest){

        String[] csvHeader = {"Client Name","Value"};
        String[] mapping = {"clientName", "value"};
        List<CompanyChartData> csvBody = clientDataObjectRequest.getCompanyChartData();

        ByteArrayInputStream byteArrayOutputStream;

        try(
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                CSVPrinter csvPrinter = new CSVPrinter(
                        new PrintWriter(out),
                        CSVFormat.DEFAULT.withHeader(csvHeader).withDelimiter(';')
                );
        ) {

            for (CompanyChartData record: csvBody) {
                String data = record.getClientName() + ", " + record.getValue();
                csvPrinter.printRecord(data.replaceAll("\"", ""));

            }
            csvPrinter.flush();
            byteArrayOutputStream = new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }

        InputStreamResource fileInputStream = new InputStreamResource(byteArrayOutputStream);

        String dateFile = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String timeFile = new SimpleDateFormat("HH-mm").format(new java.util.Date());
        String fileName = clientDataObjectRequest.getMetricName().replace(" ","_") + "_" + dateFile + "T" + timeFile;

        HttpHeaders headers = new HttpHeaders();
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName + ".csv");
        headers.set(HttpHeaders.CONTENT_TYPE, "text/csv");

        return new ResponseEntity<>(
                fileInputStream,
                headers,
                HttpStatus.OK
        );
    }

    @PostMapping(value = "/exportfile/csv", produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<InputStreamResource> exportCSV(@RequestBody ClientDataObjectRequest clientDataObjectRequest) {

        ByteArrayInputStream file = generateCSVExcelService.generateFileDetails(clientDataObjectRequest);
        //ByteArrayInputStream file = generateFileDetails(clientDataObjectRequest);

        String dateFile = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String timeFile = new SimpleDateFormat("HH-mm").format(new java.util.Date());
        String fileName = clientDataObjectRequest.getMetricName().replace(" ","_") + "_" + dateFile + "T" + timeFile;

        HttpHeaders headers = new HttpHeaders();
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName +".csv");
        headers.set(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_JSON_VALUE);

        return new ResponseEntity<>(new InputStreamResource(file), headers, HttpStatus.OK);
    }

    @PostMapping(value = "/exportfile/excel", consumes = MediaType.APPLICATION_JSON_VALUE)
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
