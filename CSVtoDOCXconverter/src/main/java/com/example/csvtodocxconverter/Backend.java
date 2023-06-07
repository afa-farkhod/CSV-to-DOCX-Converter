package com.example.csvtodocxconverter;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

class Backend {

    public String importFile(java.io.File file) {
        if (file == null) {
            return "";
        }
        return file.getAbsolutePath();
    }

    public String importImage(java.io.File file) { // << added
        if (file == null) {
            return "";
        }
        return file.getAbsolutePath();
    }

    public String convertToDOCX(String csvFilePath) {
        if (csvFilePath.isEmpty()) {
            return "No CSV file selected";
        }

        // Prepare the output DOCX file path
        String docxFilePath = csvFilePath.replace(".csv", ".docx");

        try (BufferedReader br = new BufferedReader(new FileReader(csvFilePath));
             XWPFDocument doc = new XWPFDocument()) {

            XWPFTable table = doc.createTable();
            String line;
            while ((line = br.readLine()) != null) {
                String[] values = line.split(",");
                XWPFTableRow row = table.createRow();
                for (String value : values) {
                    row.createCell().setText(value);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(docxFilePath)) {
                doc.write(fos);
            }

            return docxFilePath;

        } catch (IOException e) {
            e.printStackTrace();
            return "Error occurred during conversion";
        }

    }
}
