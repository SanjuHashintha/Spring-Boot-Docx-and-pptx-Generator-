package com.example.docx;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class DocxManipulator {

    public void manipulateDocxFile(String inputFilePath, String outputFilePath, List<Integer> data) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // Assuming you have only one table in the DOCX file
            XWPFTable table = document.getTables().get(0);

            // Assuming the table has 2 columns, you can adjust the logic if needed
            List<XWPFTableRow> rows = table.getRows();

            // Start from the second row (skip the header row)
            for (int i = 1; i < rows.size(); i++) {
                XWPFTableRow row = rows.get(i);
                List<XWPFTableCell> cells = row.getTableCells();

                // Assuming the second column (index 1) is where you want to insert the data
                if (i - 1 < data.size()) {
                    XWPFTableCell cell = cells.get(1);
                    cell.setText(String.valueOf(data.get(i - 1)));
                }
            }
            // Generate the pie chart
            DefaultPieDataset pieDataset = new DefaultPieDataset();
            DefaultCategoryDataset barDataset = new DefaultCategoryDataset();

            for (int i = 1; i < rows.size(); i++) {
                pieDataset.setValue(rows.get(i).getCell(0).getText(), Integer.parseInt(rows.get(i).getCell(1).getText()));
                barDataset.addValue(Integer.parseInt(rows.get(i).getCell(1).getText()), "Count", rows.get(i).getCell(0).getText());
            }

            JFreeChart pieChart = ChartFactory.createPieChart("Gender Distribution (Pie Chart)", pieDataset, true, true, false);
            JFreeChart barChart = ChartFactory.createBarChart("Gender Distribution (Bar Chart)", "Gender", "Count", barDataset);

            // Create byte array for chart images
            byte[] pieChartImageBytes = ChartUtils.encodeAsPNG(pieChart.createBufferedImage(500, 300));
            byte[] barChartImageBytes = ChartUtils.encodeAsPNG(barChart.createBufferedImage(500, 300));

            // Add the charts as images to the document
            addChartToDocument(document, pieChartImageBytes, "Pie Chart");
            addChartToDocument(document, barChartImageBytes, "Bar Chart");


            // Save the updated document to a new file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                document.write(fos);
            }
        }
    }

    private static void addChartToDocument(XWPFDocument doc, byte[] chartImageBytes, String chartTitle) throws IOException {
        // Create paragraph for the chart title
        XWPFParagraph titleParagraph = doc.createParagraph();
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText(chartTitle);
        titleRun.setBold(true);
        titleRun.setFontSize(12);

        // Add the chart image to the document
        XWPFParagraph paragraph = doc.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        run.addBreak();
        try (ByteArrayInputStream chartImageStream = new ByteArrayInputStream(chartImageBytes)) {
            run.addPicture(chartImageStream, Document.PICTURE_TYPE_PNG, "chart.png", Units.toEMU(500), Units.toEMU(300));
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

}
