package com.example.docx;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.apache.poi.sl.usermodel.PictureData.PictureType;
import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class PpxManipulator {
    public void manipulatePpxFile(String inputFilePath, String outputFilePath, List<Integer> data) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {

            // Assuming you have only one slide in the PPTX file
            XSLFSlide slide = ppt.getSlides().get(0);

            // Get the shapes from the slide
            List<XSLFShape> shapes = slide.getShapes();

            // Generate the pie chart
            DefaultPieDataset pieDataset = new DefaultPieDataset();
            DefaultCategoryDataset barDataset = new DefaultCategoryDataset();

            // Find the table shape (if there's a table on the slide)
            for (XSLFShape shape : shapes) {
                if (shape instanceof XSLFTable) {
                    XSLFTable table = (XSLFTable) shape;

                    // Assuming the table has 2 columns, you can adjust the logic if needed

                    List<XSLFTableRow> rows = table.getRows();
                    int rowIndex = 0;

                    // Start from the second row (skip the header row)
                    for (Integer rowData : data) {
                        if (rowIndex + 1 < rows.size()) {
                            XSLFTableRow row = rows.get(rowIndex + 1); // Skip the header row
                            List<XSLFTableCell> cells = row.getCells();

                            // Assuming the second column (index 1) is where you want to insert the data
                            if (cells.size() > 1) {
                                XSLFTableCell cell = cells.get(1);
                                cell.setText(String.valueOf(rowData));
                                pieDataset.setValue(cells.get(0).getText(), rowData);
                                barDataset.addValue(rowData, "Count",cells.get(0).getText() );
                            }
                        }
                        rowIndex++;
                    }
                }
            }


            JFreeChart pieChart = ChartFactory.createPieChart("Gender Distribution (Pie Chart)", pieDataset, true, true, false);
            JFreeChart barChart = ChartFactory.createBarChart("Gender Distribution (Bar Chart)", "Gender", "Count", barDataset);

            // Create byte array for chart images
            byte[] pieChartImageBytes = ChartUtils.encodeAsPNG(pieChart.createBufferedImage(500, 300));
            byte[] barChartImageBytes = ChartUtils.encodeAsPNG(barChart.createBufferedImage(500, 300));

            // Add the charts as images to the document
            addChartToSlideShow(ppt, pieChartImageBytes, "Pie Chart");
            addChartToSlideShow(ppt, barChartImageBytes, "Bar Chart");


            // Save the updated document to a new file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                ppt.write(fos);
            }
        } catch (XmlException e) {
            throw new RuntimeException(e);
        }
    }

    private static void addChartToSlideShow(
            XMLSlideShow slideShow, byte[] chartImageBytes, String chartTitle) throws IOException, XmlException {
        // Create a new slide
        XSLFSlide slide = slideShow.createSlide();

        // Create a paragraph for the chart title
        XSLFTextShape titleShape = slide.createTextBox();
        titleShape.setAnchor(new java.awt.Rectangle(50, 50, 600, 50)); // You can adjust the position and size as needed
        XSLFTextParagraph titleParagraph = titleShape.addNewTextParagraph();
        titleParagraph.setTextAlign(TextParagraph.TextAlign.CENTER);
        XSLFTextRun titleRun = titleParagraph.addNewTextRun();
        titleRun.setText(chartTitle);
        titleRun.setBold(true);
        titleRun.setFontSize(12.0);

        // Add the chart image to the slide
        try (ByteArrayInputStream chartImageStream = new ByteArrayInputStream(chartImageBytes)) {
            PictureData chartPictureData = slideShow.addPicture(chartImageStream, PictureType.PNG);
            XSLFPictureShape chartPictureShape = slide.createPicture(chartPictureData);
            chartPictureShape.setAnchor(new java.awt.Rectangle(50, 100, 500, 300)); // You can adjust the position and size as needed
        }
    }
}
