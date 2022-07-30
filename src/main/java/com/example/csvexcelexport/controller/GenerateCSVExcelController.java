package com.example.csvexcelexport.controller;


import com.example.csvexcelexport.pojo.ClientDataObjectRequest;
import com.example.csvexcelexport.pojo.ClientDataRequest;
import com.example.csvexcelexport.pojo.CompanyChartData;
import com.example.csvexcelexport.utils.Constants;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiResponse;
import io.swagger.annotations.ApiResponses;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xddf.usermodel.text.XDDFTextBody;
import org.apache.poi.xssf.usermodel.*;
import org.apache.tomcat.util.http.fileupload.ByteArrayOutputStream;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.supercsv.io.CsvBeanWriter;
import org.supercsv.io.ICsvBeanWriter;
import org.supercsv.prefs.CsvPreference;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.SimpleDateFormat;
import java.time.Year;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

@RestController
@RequestMapping("/generate")
@Api(value = "Generate files controller")
public class GenerateCSVExcelController {


    /**
     * Method to generate and download a csv file using commons-csv
     *
     * @author <a href="https://github.com/diegocmatte">diegocmatte</a>
     *
     * @param clientDataRequest Request body which will fill the file with the data.
     * @param metricName Request parameter which will show what the metric are being generated.
     * @param toggle Not required parameter which shows what kind of data is.
     * @return Downloadable csv file.
     */

    @ApiOperation(value = "Generate a csv file")
    @ApiResponses(value = {
            @ApiResponse(code = 200, message = "Return a csv file")
    })
    @GetMapping(value = "/v1/csv",
                produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public ResponseEntity<InputStreamResource> exportCSV(@RequestBody List<ClientDataRequest> clientDataRequest,
                                                         @RequestParam String metricName,
                                                         @RequestParam(required = false) String toggle){

        String[] csvHeader = {"Client Name", "Value"};

        List<List<ClientDataRequest>> csvBody = new ArrayList<>();
        csvBody.add(clientDataRequest);

        ByteArrayInputStream byteArrayOutputStream;

        try(
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                // defining the CSV printer
                CSVPrinter csvPrinter = new CSVPrinter(
                        new PrintWriter(out),
                        // withHeader is optional
                        CSVFormat.DEFAULT.withHeader(csvHeader)
                );
        ) {
            // populating the CSV content
            for (List<ClientDataRequest> record : csvBody) {
                csvPrinter.printRecords(record);
            }
            // writing the underlying stream
            csvPrinter.flush();
            byteArrayOutputStream = new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }

        InputStreamResource fileInputStream = new InputStreamResource(byteArrayOutputStream);
        String timestamp = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String csvFileName;
        if(toggle != null) {
            csvFileName = metricName + "_" + toggle + "_" + timestamp + ".csv";
        } else {
            csvFileName = metricName + "_" + timestamp + ".csv";
        }

        // setting HTTP headers
        HttpHeaders headers = new HttpHeaders();
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + csvFileName);
        // defining the custom Content-Type
        headers.set(HttpHeaders.CONTENT_TYPE, "text/csv");

        return new ResponseEntity<>(
                fileInputStream,
                headers,
                HttpStatus.OK
        );

    }

    /**
     * Method to generate and download a csv file using supercsv
     *
     * @author <a href="https://github.com/diegocmatte">diegocmatte</a>
     *
     * @param clientDataRequest Request body which will fill the file with the data.
     * @param metricName Request parameter which will show what the metric are being generated.
     * @param toggle Not required parameter which shows what kind of data is.
     * @param response
     * @return
     * @throws IOException
     */
    @GetMapping(value = "/v2/csv", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public ResponseEntity<Void> exportCSV_v2(@RequestBody List<ClientDataRequest> clientDataRequest,
                                             @RequestParam String metricName,
                                             @RequestParam(required = false) String toggle,
                                             HttpServletResponse response) throws IOException {

        response.setContentType("text/csv");
        String timestamp = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String csvFileName;
        if(toggle != null) {
            csvFileName = metricName + "_" + toggle + "_" + timestamp + ".csv";
        } else {
            csvFileName = metricName + "_" + timestamp + ".csv";
        }
        ICsvBeanWriter csvWriter = new CsvBeanWriter(response.getWriter(), CsvPreference.STANDARD_PREFERENCE);
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=" + csvFileName;
        response.setHeader(headerKey, headerValue);


        String[] csvHeader = {"Client Name", "Value"};
        String[] nameMapping = {"clientName", "value"};

        csvWriter.writeHeader(csvHeader);

        for(ClientDataRequest clientData: clientDataRequest){
            csvWriter.write(clientData, nameMapping);
        }

        csvWriter.close();



        return ResponseEntity.ok().build();

    }

    @PostMapping(value = "/v3/csv/graphic", produces = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<InputStreamResource> export_V3(@RequestBody List<CompanyChartData> companyChartData,
                                                         @RequestParam String metricName) {

        ByteArrayInputStream file = generateFileDetails(companyChartData, metricName);

        HttpHeaders headers = new HttpHeaders();
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=metric.xlsx");
        headers.set(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_JSON_VALUE);

        return new ResponseEntity<>(new InputStreamResource(file), headers, HttpStatus.OK);
    }

    private ByteArrayInputStream generateFileDetails(List<CompanyChartData> companyChartData, String metricName){

        ByteArrayOutputStream file = new ByteArrayOutputStream();

        try{
            try(XSSFWorkbook workbook = new XSSFWorkbook()){
                XSSFSheet sheet = workbook.createSheet(metricName);
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerFont.setUnderline(Font.U_SINGLE);
                CellStyle headerCellStyle = workbook.createCellStyle();
                headerCellStyle.setFillBackgroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
                headerCellStyle.setFont(headerFont);

                XSSFDrawing drawing = sheet.createDrawingPatriarch();
                XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 2, 1, 11, 10 + companyChartData.size());

                Row headerRow = sheet.createRow(anchor.getRow2()+1);
                headerRow.createCell(0).setCellValue("Client Name");
                headerRow.createCell(1).setCellValue("Value");

                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    headerRow.getCell(i).setCellStyle(headerCellStyle);
                }

                for (int i = 0; i < companyChartData.size(); i++) {
                    Row row = sheet.createRow(anchor.getRow2() + 2 + i);
                    row.createCell(0).setCellValue(companyChartData.get(i).getClientName());
                    sheet.autoSizeColumn(0);
                    row.createCell(1).setCellValue(companyChartData.get(i).getValue());
                    sheet.autoSizeColumn(1);
                }

                XSSFChart chart = drawing.createChart(anchor);
                chart.setTitleText(metricName);
                chart.setTitleOverlay(false);

                //XDDFChartLegend legend = chart.getOrAddLegend();
                //legend.setPosition(LegendPosition.TOP_RIGHT);

                XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
                bottomAxis.setTitle("Client Name");
                XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
                leftAxis.setTitle("Value");
                leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

                XDDFDataSource<String> clientNames = XDDFDataSourcesFactory.fromStringCellRange(sheet,
                        new CellRangeAddress(anchor.getRow2() + 2, anchor.getRow2() + 1 + companyChartData.size(), 0, 0));

                XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                        new CellRangeAddress(anchor.getRow2() + 2, anchor.getRow2() +1 + companyChartData.size(), 1, 1));

                XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
                XDDFChartData.Series series = data.addSeries(clientNames, values);
                series.setTitle(metricName, null);
                data.setVaryColors(false);
                chart.plot(data);

                XDDFBarChartData bar = (XDDFBarChartData) data;
                bar.setBarDirection(BarDirection.BAR);

                Row bottomRow = sheet.createRow(anchor.getRow2() + companyChartData.size() + 3);
                bottomRow.createCell(0).setCellValue("some text here that is needed to be larger than 5 columns");

                workbook.write(file);
            }
        } catch (Exception e){

        }

        return new ByteArrayInputStream(file.toByteArray());

    }

    @PostMapping(value = "/v4/csv/graphic", consumes = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<InputStreamResource> export_V4(@RequestBody ClientDataObjectRequest clientDataObjectRequest) {

        ByteArrayInputStream file = generateFileDetails_v2(clientDataObjectRequest);

        String dateFile = new SimpleDateFormat("MM-dd-yyyy").format(new java.util.Date());
        String timeFile = new SimpleDateFormat("HH-mm").format(new java.util.Date());
        String fileName = clientDataObjectRequest.getMetricName().replace(" ","_") + "_" + dateFile + "T" + timeFile;

        HttpHeaders headers = new HttpHeaders();
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName + ".xlsx");
        headers.set(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_JSON_VALUE);

        return new ResponseEntity<>(new InputStreamResource(file), headers, HttpStatus.OK);
    }

    private ByteArrayInputStream generateFileDetails_v2(ClientDataObjectRequest clientDataObjectRequest) {

        ByteArrayOutputStream file = new ByteArrayOutputStream();
        String dataFormat = clientDataObjectRequest.getDataFormatCodeValue();

        try{
            try(XSSFWorkbook workbook = new XSSFWorkbook()){
                XSSFSheet sheet = workbook.createSheet(clientDataObjectRequest.getMetricName());
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerFont.setUnderline(Font.U_SINGLE);
                CellStyle headerCellStyle = workbook.createCellStyle();
                headerCellStyle.setFillBackgroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
                headerCellStyle.setFont(headerFont);

                XSSFDrawing drawing = sheet.createDrawingPatriarch();
                XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 2, 1, 11, 10 + clientDataObjectRequest.getCompanyChartData().size());

                Row headerRow = sheet.createRow(anchor.getRow2() + 1);
                headerRow.createCell(0).setCellValue("Client Name");
                headerRow.createCell(1).setCellValue("Value ("+clientDataObjectRequest.getDataFormatCodeValue()+")");

                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    headerRow.getCell(i).setCellStyle(headerCellStyle);
                }

                for (int i = 0; i < clientDataObjectRequest.getCompanyChartData().size(); i++) {
                    Row row = sheet.createRow(anchor.getRow2() + 2 + i);
                    row.createCell(0).setCellValue(clientDataObjectRequest.getCompanyChartData().get(i).getClientName());
                    row.createCell(1).setCellValue(clientDataObjectRequest.getCompanyChartData().get(i).getValue());
                }

                XSSFChart chart = drawing.createChart(anchor);
                chart.setTitleText(clientDataObjectRequest.getMetricName());
                chart.getCTChart().getTitle().getTx().getRich().getPArray(0).getRArray(0).getRPr().setSz(2000);
                chart.setTitleOverlay(false);

                //XDDFChartLegend legend = chart.getOrAddLegend();
                //legend.setPosition(LegendPosition.TOP_RIGHT);

                XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
                bottomAxis.setTitle("Client Name");

                XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
                leftAxis.setTitle("Value ("+clientDataObjectRequest.getDataFormatCodeValue()+")");
                leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

                sheet.shiftColumns(0,1,2);
                sheet.autoSizeColumn(2);
                sheet.autoSizeColumn(3);

                XDDFDataSource<String> clientNames = XDDFDataSourcesFactory.fromStringCellRange(sheet,
                        new CellRangeAddress(anchor.getRow2() + 2, anchor.getRow2() + 1 + clientDataObjectRequest.getCompanyChartData().size(), 2, 2));

                XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                        new CellRangeAddress(anchor.getRow2() + 2, anchor.getRow2() +1 + clientDataObjectRequest.getCompanyChartData().size(), 3, 3));

                XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
                XDDFChartData.Series series = data.addSeries(clientNames, values);
                series.setTitle(clientDataObjectRequest.getMetricName(), null);
                data.setVaryColors(false);
                chart.plot(data);

                XDDFBarChartData bar = (XDDFBarChartData) data;
                bar.setBarDirection(BarDirection.BAR);

                Row bottomRow = sheet.createRow(anchor.getRow2() + clientDataObjectRequest.getCompanyChartData().size() + 3);
                bottomRow.createCell(0).setCellValue(Constants.COPYRIGHT_FOOTER);

                workbook.write(file);
            }
        } catch (Exception e){

        }

        return new ByteArrayInputStream(file.toByteArray());
    }

}
