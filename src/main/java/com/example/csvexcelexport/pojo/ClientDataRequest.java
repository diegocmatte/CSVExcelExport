package com.example.csvexcelexport.pojo;

import com.fasterxml.jackson.annotation.JsonProperty;

public class ClientDataRequest {

    @JsonProperty("clientName")
    private String clientName;
    @JsonProperty("value")
    private Object value;

    @Override
    public String toString() {
        final StringBuilder sb = new StringBuilder();
        sb.append(clientName).append(',').append(value);
        return sb.toString();
    }
}
