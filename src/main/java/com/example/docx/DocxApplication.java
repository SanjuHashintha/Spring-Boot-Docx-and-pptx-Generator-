package com.example.docx;

import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.SpringApplication;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class DocxApplication {

	public static void main(String[] args) {
		SpringApplication.run(DocxApplication.class, args);
		// Replace these file paths with your input and output file paths
		String inputFilePath = "/home/hsenid/Documents/docx/src/main/resources/input.docx";
		String outputFilePath = "output.docx";

		String inputFilePathPptx = "/home/hsenid/Documents/docx/src/main/resources/input.pptx";
		String outputFilePathPptx= "output.pptx";

		// Example data to fill the second column of the table
		List<Integer> data = new ArrayList<>();
		data.add(10);
		data.add(7);
		data.add(9);
		data.add(6);
		// Add more data as needed

		DocxManipulator docxManipulator = new DocxManipulator();
		try {
			docxManipulator.manipulateDocxFile(inputFilePath, outputFilePath, data);
			System.out.println("New DOCX file generated successfully.");
		} catch (IOException e) {
			System.err.println("Error manipulating the DOCX file: " + e.getMessage());
		}

		PpxManipulator ppxManipulator = new PpxManipulator();
		try {
			ppxManipulator.manipulatePpxFile(inputFilePathPptx, outputFilePathPptx, data);
			System.out.println("New PPTX file generated successfully.");
		} catch (IOException e) {
			System.err.println("Error manipulating the PPTX file: " + e.getMessage());
		}

	}

}
