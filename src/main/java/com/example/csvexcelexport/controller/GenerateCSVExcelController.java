package com.example.csvexcelexport.controller;

import com.example.csvexcelexport.pojo.ClientDataRequest;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiResponse;
import io.swagger.annotations.ApiResponses;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
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
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
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

}
