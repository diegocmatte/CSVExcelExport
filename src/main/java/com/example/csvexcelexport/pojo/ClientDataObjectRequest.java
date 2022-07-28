package com.example.csvexcelexport.pojo;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

import java.util.List;

@Data
public class ClientDataObjectRequest {

    @JsonProperty("metricName")
    private String metricName;

    @JsonProperty("dataFormatCodeValue")
    private String dataFormatCodeValue;

    @JsonProperty("clientDataRequest")
    private List<CompanyChartData> companyChartData;

}
