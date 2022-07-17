package com.example.csvexcelexport.pojo;

import lombok.Data;

@Data
public class ClientDataRequest {

    private String clientName;
    private String value;

    @Override
    public String toString() {
        final StringBuilder sb = new StringBuilder();
        sb.append(clientName).append(',').append(value);
        return sb.toString();
    }
}
