package com.maruti.parking.common.util;

import org.apache.commons.math3.util.Precision;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.xmlbeans.SchemaTypeSystem;
import org.apache.xmlbeans.impl.schema.SchemaTypeSystemImpl;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.springframework.boot.context.properties.bind.DefaultValue;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Date;
import java.util.List;

@Component
public class ExcelUtil {

    public Sheet createHeading(Sheet sheet, String... headings) {

        Row header = sheet.createRow(0);
        int col = 0;
        for (String heading : headings) {
            Cell cell = header.createCell(col);
            cell.setCellValue(heading);
            col++;
        }
        return sheet;
    }



    public Workbook fillDataInSheet(List<Object> data1, String[] headers, String sheetName, XSSFWorkbook workbook) {

        Sheet sheet = workbook.createSheet(sheetName);
        Row headerrow = sheet.createRow(0);
        Field[] fields = data1.get(0).getClass().getDeclaredFields();
        createHeading(sheet, headers);
        Row rowHeader = sheet.getRow(0);
        styleHeader(rowHeader, workbook);

        // Create a cell style for data cells with borders
        XSSFCellStyle dataCellStyle = workbook.createCellStyle();
        dataCellStyle.setBorderTop(BorderStyle.THIN);
        dataCellStyle.setBorderBottom(BorderStyle.THIN);
        dataCellStyle.setBorderLeft(BorderStyle.THIN);
        dataCellStyle.setBorderRight(BorderStyle.THIN);

        int row = 1;
        int col;
        for (Object item : data1) {
            Row rows = sheet.createRow(row++);
            col = 0;
            for (Field field : fields) {
                field.setAccessible(true);
                try {
                    Object item1 = field.get(item);
                    Cell cell = rows.createCell(col++);

                    if (item1 instanceof String) {
                        cell.setCellValue((String) item1);
                    } else if (item1 instanceof Number) {
                        cell.setCellValue(((Number) item1).doubleValue());
                    } else if (item1 instanceof Boolean) {
                        cell.setCellValue((Boolean) item1);
                    } else if (item1 instanceof Date) {
                        cell.setCellValue((Date) item1);
                    } else if (item1 instanceof List) {
                        col--;
                        for (Object object : (List) item1) {
                            Cell cell1 = rows.createCell(col++);
                            cell1.setCellValue((String) object);
                            cell1.setCellStyle(dataCellStyle);
                        }
                    }
                    cell.setCellStyle(dataCellStyle);
                } catch (IllegalAccessException e) {
                    throw new RuntimeException("Failed to access fieled" + field.getName(), e);
                }
            }

        }

        for (int i = 0; i < rowHeader.getLastCellNum(); i++) {
            sheet.autoSizeColumn(i);
        }

        return workbook;
    }

    public Workbook styleHeader(Row headerRow, XSSFWorkbook workbook) {
        // Create cell style for headers with background color and borders
        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.ROYAL_BLUE.index);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Sheet sheet = headerRow.getSheet();
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell headerCell = headerRow.getCell(i);
            if (headerCell != null) {
                headerCell.setCellStyle(headerStyle);
            }
            sheet.autoSizeColumn(i);
        }
        // Auto-size columns to fit content


        return workbook;
    }


    public void createPieChart(XSSFSheet sheet, String tittleTxt, int categoryColumn, int rangeColumn) {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 4, 7, 20);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(tittleTxt);
        chart.setTitleOverlay(false);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);
        // Define the cell range for the data source and customize the legend colors
        CellRangeAddress categoryRange = new CellRangeAddress(1, sheet.getLastRowNum(), categoryColumn, categoryColumn);
        // Change this to your desired range for categories
        CellRangeAddress valueRange = new CellRangeAddress(1, sheet.getLastRowNum(), rangeColumn, rangeColumn);
        // Change this to your desired range for values
        XDDFDataSource<String> sourceSystem = XDDFDataSourcesFactory.fromStringCellRange(sheet, categoryRange);
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, valueRange);
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        data.setVaryColors(true);
        data.addSeries(sourceSystem, values);
        chart.plot(data);

        if (chart.getCTChart().getAutoTitleDeleted() == null) chart.getCTChart().addNewAutoTitleDeleted();
        chart.getCTChart().getAutoTitleDeleted().setVal(false);

    }

    public void createColumnChart(XSSFSheet sheet, @DefaultValue("Title") String chartTitle, @DefaultValue("Category") String categoryName, int[] rowRange, @DefaultValue("1") int categoryColumn, @DefaultValue("2") int valueColumn) {

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 5, 2, 15, 20);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(chartTitle);
        CTChart ctChart = ((XSSFChart) chart).getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
        CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
        ctBoolean.setVal(false);
        ctBarChart.addNewBarDir().setVal(STBarDir.COL);

        // Series
        CTBarSer ctBarSer = ctBarChart.addNewSer();
        CTDLbls ctDLbls = ctBarSer.addNewDLbls();
        // Show value
        ctDLbls.addNewShowVal().setVal(true);
        // Hide category name
        ctDLbls.addNewShowCatName().setVal(false);
        // Hide series name
        ctDLbls.addNewShowSerName().setVal(false);
        ctDLbls.addNewShowLegendKey().setVal(false);
        ctBarSer.addNewIdx().setVal(0);
        CTSerTx ctSerTx = ctBarSer.addNewTx();
        CTStrRef ctStrRef = ctSerTx.addNewStrRef();
        System.out.println(generateCellRange(sheet.getSheetName(), rowRange[0], rowRange[1], categoryColumn, categoryColumn) + "\n" +

                generateCellRange(sheet.getSheetName(), rowRange[0], rowRange[1], categoryColumn, categoryColumn));

        generateCellRange(sheet.getSheetName(), rowRange[0], rowRange[1], valueColumn, valueColumn);
        ctStrRef.setF(generateCellRange(sheet.getSheetName(), rowRange[0], rowRange[1], categoryColumn, categoryColumn));
        CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
        ctStrRef = cttAxDataSource.addNewStrRef();
        ctStrRef.setF(generateCellRange(sheet.getSheetName(), rowRange[0], rowRange[1], categoryColumn, categoryColumn));
        CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
        CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
        ctNumRef.setF(generateCellRange(sheet.getSheetName(), rowRange[0], rowRange[1], valueColumn, valueColumn));

        // Customize appearance
        ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0, 0, 0});

        // Axes
        ctBarChart.addNewAxId().setVal(123456);
        ctBarChart.addNewAxId().setVal(123457);

        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123456);
        ctCatAx.addNewScaling().addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123457);
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(123457);
        ctValAx.addNewScaling().addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.L);
        ctValAx.addNewCrossAx().setVal(123456);
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
        CTTitle ctTitle = ctCatAx.addNewTitle();
        // Set the text
        ctTitle.addNewOverlay().setVal(false);
        CTTx tx = ctTitle.addNewTx();
        CTTextBody rich = tx.addNewRich();
        rich.addNewBodyPr(); // body properties must exist, but can be empty
        CTTextParagraph para = rich.addNewP();
        CTRegularTextRun rxt = para.addNewR();
        rxt.setT(categoryName);

    }


    public String generateCellRange(String sheetName, int startRow, int endRow, int startColumn, int endColumn) {
        char startColumnChar = (char) ('A' + startColumn - 1);
        char endColumnChar = (char) ('A' + endColumn - 1);

        // Generating the Excel-style cell range format
        //for example "CitywiseData!$A$2:$A$7"
        return sheetName + "!$" + startColumnChar + "$" + startRow + ":$" + endColumnChar + "$" + endRow;
    }


}
