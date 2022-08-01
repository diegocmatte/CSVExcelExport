package com.example.csvexcelexport.utils;

import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;

public final class ChartUtils {

    public static XDDFTitle getOrSetAxisTitle(XDDFValueAxis axis) {
        try {
            java.lang.reflect.Field _ctValAx = XDDFValueAxis.class.getDeclaredField("ctValAx");
            _ctValAx.setAccessible(true);
            org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx ctValAx =
                    (org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx)_ctValAx.get(axis);
            if (!ctValAx.isSetTitle()) {
                ctValAx.addNewTitle();
            }
            XDDFTitle title = new XDDFTitle(null, ctValAx.getTitle());
            return title;
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    public static XDDFTitle getOrSetAxisTitle(XDDFCategoryAxis axis) {
        try {
            java.lang.reflect.Field _ctCatAx = XDDFCategoryAxis.class.getDeclaredField("ctCatAx");
            _ctCatAx.setAccessible(true);
            org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx ctCatAx =
                    (org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx)_ctCatAx.get(axis);
            if (!ctCatAx.isSetTitle()) {
                ctCatAx.addNewTitle();
            }
            XDDFTitle title = new XDDFTitle(null, ctCatAx.getTitle());
            return title;
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    public static void solidFillSeries(XDDFChartData data, int index, PresetColor color) {
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(color));
        XDDFChartData.Series series = data.getSeries(index);
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series.setShapeProperties(properties);
    }

    public static void setDataLabels(XDDFChartData.Series series, int pos, boolean... show) {
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
