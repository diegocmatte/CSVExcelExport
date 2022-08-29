package com.example.csvexcelexport.service;

import com.example.csvexcelexport.pojo.ClientDataObjectRequest;

import com.example.csvexcelexport.pojo.CompanyChartData;
import com.example.csvexcelexport.utils.Constants;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.QuoteMode;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.tomcat.util.http.fileupload.ByteArrayOutputStream;
import org.springframework.core.io.InputStreamResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.List;

@Service
public class GenerateCSVExcelService {

    public InputStreamResource generateCSV(ClientDataObjectRequest clientDataObjectRequest) {

        String[] csvHeader = {"Client Name","Value"};
        List<CompanyChartData> csvBody = clientDataObjectRequest.getCompanyChartData();

        ByteArrayInputStream byteArrayOutputStream;

        try(
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                CSVPrinter csvPrinter = new CSVPrinter(
                        new PrintWriter(out),
                        CSVFormat.DEFAULT.withHeader(csvHeader).withQuoteMode(QuoteMode.NON_NUMERIC)
                );
        ) {

            for (CompanyChartData record: csvBody) {
                csvPrinter.printRecord(record.getClientName(), record.getValue());

            }
            csvPrinter.flush();
            byteArrayOutputStream = new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }

        return new InputStreamResource(byteArrayOutputStream);

    }

    public ByteArrayInputStream generateFileDetailsWithChart(ClientDataObjectRequest clientDataObjectRequest) {

        ByteArrayOutputStream file = new ByteArrayOutputStream();

        try{
            try(XSSFWorkbook workbook = new XSSFWorkbook()){
                XSSFSheet sheet = workbook.createSheet(clientDataObjectRequest.getMetricName().replace("/","_"));
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerFont.setUnderline(Font.U_SINGLE);
                CellStyle headerCellStyle = workbook.createCellStyle();
                headerCellStyle.setFillBackgroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
                headerCellStyle.setFont(headerFont);

                XSSFDrawing drawing = sheet.createDrawingPatriarch();
                XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 2, 1, 15, 15 + clientDataObjectRequest.getCompanyChartData().size());

                Row headerRow = sheet.createRow(anchor.getRow2() + 1);
                headerRow.createCell(0).setCellValue(Constants.CLIENT_NAME);
                headerRow.createCell(1).setCellValue("Value ("+clientDataObjectRequest.getDataFormatCodeValue()+")");

                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    headerRow.getCell(i).setCellStyle(headerCellStyle);
                }

                CellStyle rowCellStyle = workbook.createCellStyle();

                if (clientDataObjectRequest.getDataFormatCodeValue().equalsIgnoreCase(Constants.DATA_FORMAT_PERCENTAGE)) {
                    rowCellStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
                    for (int i = 0; i < clientDataObjectRequest.getCompanyChartData().size(); i++) {
                        Row row = sheet.createRow(anchor.getRow2() + 2 + i);
                        row.createCell(0).setCellValue(clientDataObjectRequest.getCompanyChartData().get(i).getClientName());
                        row.createCell(1).setCellValue(clientDataObjectRequest.getCompanyChartData().get(i).getValue()/100);
                        //rowCellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00%"));
                        row.getCell(1).setCellStyle(rowCellStyle);
                    }
                }

                if (clientDataObjectRequest.getDataFormatCodeValue().equalsIgnoreCase(Constants.DATA_FORMAT_CURRENCY)) {
                    rowCellStyle.setDataFormat(workbook.createDataFormat().getFormat("$0.00"));
                    for (int i = 0; i < clientDataObjectRequest.getCompanyChartData().size(); i++) {
                        Row row = sheet.createRow(anchor.getRow2() + 2 + i);
                        row.createCell(0).setCellValue(clientDataObjectRequest.getCompanyChartData().get(i).getClientName());
                        row.createCell(1).setCellValue(clientDataObjectRequest.getCompanyChartData().get(i).getValue());
                        row.getCell(1).setCellStyle(rowCellStyle);
                    }
                }

                XSSFChart chart = drawing.createChart(anchor);
                chart.setTitleText(clientDataObjectRequest.getMetricName());
                chart.getCTChart().getTitle().getTx().getRich().getPArray(0).getRArray(0).getRPr().setSz(2000);
                chart.setTitleOverlay(false);

                //XDDFChartLegend legend = chart.getOrAddLegend();
                //legend.setPosition(LegendPosition.TOP_RIGHT);

                XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
                XDDFTitle title = getOrSetAxisTitle(bottomAxis);
                title.setOverlay(false);
                title.setText(Constants.CLIENT_NAME);
                title.getBody().getParagraph(0).addDefaultRunProperties().setFontSize(12d);


                XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
                title = getOrSetAxisTitle(leftAxis);
                title.setOverlay(false);
                title.setText("Value ("+clientDataObjectRequest.getDataFormatCodeValue()+")");
                title.getBody().getParagraph(0).addDefaultRunProperties().setFontSize(12d);

                if(clientDataObjectRequest.getDataFormatCodeValue().equalsIgnoreCase(Constants.DATA_FORMAT_PERCENTAGE)) {
                    leftAxis.setMaximum(1.0);
                }
                leftAxis.setCrosses(AxisCrosses.MAX);
                leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

                sheet.shiftColumns(0,1,2);
                sheet.autoSizeColumn(2);
                sheet.autoSizeColumn(3);


                XDDFDataSource<String> clientNames = XDDFDataSourcesFactory.fromStringCellRange(sheet,
                        new CellRangeAddress(anchor.getRow2() + 2,anchor.getRow2() + 1 + clientDataObjectRequest.getCompanyChartData().size(),2,2));

                XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                        new CellRangeAddress(anchor.getRow2() + 2,anchor.getRow2() + 1 + clientDataObjectRequest.getCompanyChartData().size(),3,3));

                XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
                XDDFChartData.Series series = data.addSeries(clientNames, values);
                series.setTitle(clientDataObjectRequest.getMetricName(), null);

                setDataLabels(series,7,true); // pos 7 = INT_OUT_END, showVal = true

                chart.plot(data);

                XDDFBarChartData bar = (XDDFBarChartData) data;
                bar.setBarDirection(BarDirection.BAR);

                bottomAxis.setOrientation(AxisOrientation.MAX_MIN);


                solidFillSeries(data, 0, PresetColor.BLUE);

                Row bottomRow = sheet.createRow(anchor.getRow2() + clientDataObjectRequest.getCompanyChartData().size() + 3);
                bottomRow.createCell(0).setCellValue(Constants.COPYRIGHT_FOOTER);

                workbook.write(file);
            }
        } catch (Exception e){

        }

        return new ByteArrayInputStream(file.toByteArray());
    }

    private static XDDFTitle getOrSetAxisTitle(XDDFValueAxis axis) {
        try {
            java.lang.reflect.Field _ctValAx = XDDFValueAxis.class.getDeclaredField("ctValAx");
            _ctValAx.setAccessible(true);
            org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx ctValAx =
                    (org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx)_ctValAx.get(axis);
            if (!ctValAx.isSetTitle()) {
                ctValAx.addNewTitle();
            }
            return new XDDFTitle(null, ctValAx.getTitle());
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    private static XDDFTitle getOrSetAxisTitle(XDDFCategoryAxis axis) {
        try {
            java.lang.reflect.Field _ctCatAx = XDDFCategoryAxis.class.getDeclaredField("ctCatAx");
            _ctCatAx.setAccessible(true);
            org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx ctCatAx =
                    (org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx)_ctCatAx.get(axis);
            if (!ctCatAx.isSetTitle()) {
                ctCatAx.addNewTitle();
            }
            return new XDDFTitle(null, ctCatAx.getTitle());
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    private static void solidFillSeries(XDDFChartData data, int index, PresetColor color) {
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(color));
        XDDFChartData.Series series = data.getSeries(index);
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series.setShapeProperties(properties);
    }

    private static void setDataLabels(XDDFChartData.Series series, int pos, boolean... show) {
        /*
        INT_BEST_FIT   1
        INT_B          2
        INT_CTR        3
        INT_IN_BASE    4
        INT_IN_END     5
        INT_L          6
        INT_OUT_END    7
        INT_R          8
        INT_T          9
        */

        try {
            org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls ctDLbls = null;
            if (series instanceof XDDFBarChartData.Series) {
                java.lang.reflect.Field _ctBarSer = XDDFBarChartData.Series.class.getDeclaredField("series");
                _ctBarSer.setAccessible(true);
                org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer ctBarSer =
                        (org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer)_ctBarSer.get((XDDFBarChartData.Series)series);
                if (ctBarSer.isSetDLbls()) ctBarSer.unsetDLbls();
                ctDLbls = ctBarSer.addNewDLbls();
                if (!(pos == 3 || pos == 4 || pos == 5 || pos == 7)) pos = 3; // bar chart does not provide other pos
                ctDLbls.addNewDLblPos().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos.Enum.forInt(pos));
            } else if (series instanceof XDDFLineChartData.Series) {
                java.lang.reflect.Field _ctLineSer = XDDFLineChartData.Series.class.getDeclaredField("series");
                _ctLineSer.setAccessible(true);
                org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer ctLineSer =
                        (org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer)_ctLineSer.get((XDDFLineChartData.Series)series);
                if (ctLineSer.isSetDLbls()) ctLineSer.unsetDLbls();
                ctDLbls = ctLineSer.addNewDLbls();
                if (!(pos == 3 || pos == 6 || pos == 8 || pos == 9 || pos == 2)) pos = 3; // line chart does not provide other pos
                ctDLbls.addNewDLblPos().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos.Enum.forInt(pos));
            } else if (series instanceof XDDFPieChartData.Series) {
                java.lang.reflect.Field _ctPieSer = XDDFPieChartData.Series.class.getDeclaredField("series");
                _ctPieSer.setAccessible(true);
                org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer ctPieSer =
                        (org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer)_ctPieSer.get((XDDFPieChartData.Series)series);
                if (ctPieSer.isSetDLbls()) ctPieSer.unsetDLbls();
                ctDLbls = ctPieSer.addNewDLbls();
                if (!(pos == 3 || pos == 1 || pos == 4 || pos == 5)) pos = 3; // pie chart does not provide other pos
                ctDLbls.addNewDLblPos().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos.Enum.forInt(pos));
            }// else if ...

            if (ctDLbls != null) {
                ctDLbls.addNewShowVal().setVal((show.length>0)?show[0]:false);
                ctDLbls.addNewShowLegendKey().setVal((show.length>1)?show[1]:false);
                ctDLbls.addNewShowCatName().setVal((show.length>2)?show[2]:false);
                ctDLbls.addNewShowSerName().setVal((show.length>3)?show[3]:false);
                ctDLbls.addNewShowPercent().setVal((show.length>4)?show[4]:false);
                ctDLbls.addNewShowBubbleSize().setVal((show.length>5)?show[5]:false);
                ctDLbls.addNewShowLeaderLines().setVal((show.length>6)?show[8]:false);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}
