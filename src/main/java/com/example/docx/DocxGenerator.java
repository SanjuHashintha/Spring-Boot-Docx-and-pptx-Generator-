package com.example.docx;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class DocxGenerator {

    public static void generateDocxFile(List<Data> dataList) throws IOException {
        // Load the template DOCX file
        XWPFDocument doc = new XWPFDocument(DocxGenerator.class.getResourceAsStream("/input.docx"));
        List<Integer> valueList = new ArrayList<>();
        valueList.add(10);
        valueList.add(7);

        // Get the table from the template document
        XWPFTable table = doc.getTables().get(0);

        // Update the table data with data from the backend array
        for (int i = 0; i < dataList.size(); i++) {
            Data data = dataList.get(i);
            XWPFTableRow row = table.createRow();
            row.getCell(0).setText(data.getGender());
            row.getCell(1).setText(String.valueOf(data.getCount()));
        }

        // Generate the pie chart
        DefaultPieDataset pieDataset = new DefaultPieDataset();
        DefaultCategoryDataset barDataset = new DefaultCategoryDataset();

        for (Data data : dataList) {
            pieDataset.setValue(data.getGender(), data.getCount());
            barDataset.addValue(data.getCount(), "Count", data.getGender());
        }

        JFreeChart pieChart = ChartFactory.createPieChart("Gender Distribution (Pie Chart)", pieDataset, true, true, false);
        JFreeChart barChart = ChartFactory.createBarChart("Gender Distribution (Bar Chart)", "Gender", "Count", barDataset);

        // Create byte array for chart images
        byte[] pieChartImageBytes = ChartUtils.encodeAsPNG(pieChart.createBufferedImage(500, 300));
        byte[] barChartImageBytes = ChartUtils.encodeAsPNG(barChart.createBufferedImage(500, 300));

        // Add the charts as images to the document
        addChartToDocument(doc, pieChartImageBytes, "Pie Chart");
        addChartToDocument(doc, barChartImageBytes, "Bar Chart");

        // Save the output DOCX file
        FileOutputStream outputStream = new FileOutputStream("output.docx");
        doc.write(outputStream);

        // Clean up
        outputStream.close();
        doc.close();

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

    public static class Data {
        private String gender;
        private int count;

        public Data(String gender, int count) {
            this.gender = gender;
            this.count = count;
        }

        public String getGender() {
            return gender;
        }

        public int getCount() {
            return count;
        }
    }
}
