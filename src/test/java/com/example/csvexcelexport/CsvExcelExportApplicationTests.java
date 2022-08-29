package com.example.csvexcelexport;

import com.example.csvexcelexport.controller.GenerateCSVExcelController;
import com.example.csvexcelexport.exception.ServiceExcepetion;
import com.example.csvexcelexport.pojo.ClientDataObjectRequest;
import com.example.csvexcelexport.service.GenerateCSVExcelService;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.web.client.RestTemplateBuilder;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.test.util.ReflectionTestUtils;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MockMvcBuilder;
import org.springframework.test.web.servlet.MvcResult;
import org.springframework.test.web.servlet.RequestBuilder;
import org.springframework.test.web.servlet.request.MockMvcRequestBuilders;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;

import java.io.ByteArrayInputStream;

@SpringBootTest
class CsvExcelExportApplicationTests {

    private MockMvc mockMvc;

    @Mock
    private GenerateCSVExcelService generateCSVExcelService;

    @BeforeEach
    public void init() {
        final GenerateCSVExcelController actualObject = new GenerateCSVExcelController();
        ReflectionTestUtils.setField(actualObject,"generateCSVExcelService", generateCSVExcelService);
        mockMvc = MockMvcBuilders.standaloneSetup(actualObject).build();
    }

    @Test
    @DisplayName("test 200 for /generate/exportfile/csv")
    public void returning200(){

       Mockito.when(generateCSVExcelService.generateCSV(createObject()))
               .thenReturn(this.generateCSV());

        try {
            RequestBuilder requestBuilder = MockMvcRequestBuilders.post("/generate/exportfile/csv")
                    .contentType(MediaType.APPLICATION_JSON_VALUE)
                    .accept("text/csv")
                    .content(jsonMock());
            MvcResult result = mockMvc.perform(requestBuilder).andReturn();
            Assertions.assertEquals(HttpStatus.OK.value(), result.getResponse().getStatus());
        } catch (Exception exception){
            throw new ServiceExcepetion(exception);
        }
    }

    private InputStreamResource generateCSV(){
        ClientDataObjectRequest objectRequest = createObject();
        return generateCSVExcelService.generateCSV(objectRequest);
    }

    private ClientDataObjectRequest createObject(){
        String json = jsonMock();
        return new Gson().fromJson(json, ClientDataObjectRequest.class);
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
