package com.example.csvexcelexport.pojo;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

@Data
public class CompanyChartData {

    @JsonProperty("clientName")
    private String clientName;
    @JsonProperty("value")
    private double value;

    @Override
    public String toString() {
        final StringBuilder sb = new StringBuilder();
        sb.append(clientName).append(',').append(value);
        return sb.toString();
    }
}
