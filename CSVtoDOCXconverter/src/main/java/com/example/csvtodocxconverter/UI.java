package com.example.csvtodocxconverter;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Hyperlink;
import javafx.scene.control.Label;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;


public class UI extends Application{

    private String[] paths;

    public UI() {

    }

    private Stage primaryStage;
    private FileChooser fileChooser;
    private Label fileLabel;
    private Button importButton;
    private Button importImageButton; // << added
    private Button convertButton;
    private Button clearButton;
    private Button exitButton;
    private Hyperlink downloadLink;
    private Backend backend;
    private List<String> convertedDocxFilePaths;

    public UI(Stage primaryStage, Backend backend) {
        this.primaryStage = primaryStage;
        this.backend = backend;

        fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("All Files", "*.*"));
        new FileChooser.ExtensionFilter("CSV Files", "*.csv");
        new FileChooser.ExtensionFilter("Image Files", "*.png", "*.jpg", "*.gif");

        fileLabel = new Label();
        importButton = new Button("Import File");
        importButton.setOnAction(e -> importFile());

        importImageButton = new Button("Import Image"); // << added
        importImageButton.setOnAction(e -> importImage()); // << added

        convertButton = new Button("Convert to DOCX");
        convertButton.setOnAction(e -> convertToDOCX(fileLabel.getText()));

        downloadLink = new Hyperlink();
        downloadLink.setVisible(false);

        clearButton = new Button("CLEAR");
        clearButton.setOnAction(e -> clearConvertedFiles());

        exitButton = new Button("Exit");
        exitButton.setOnAction(e -> primaryStage.close());

        convertedDocxFilePaths = new ArrayList<>();
    }

    public void show() {
        VBox vbox = new VBox(10);
        vbox.setPadding(new Insets(10));
        vbox.getChildren().addAll(fileLabel, importButton, importImageButton, convertButton, downloadLink, clearButton, exitButton);

        primaryStage.setScene(new Scene(vbox, 400, 250)); // << changed
        primaryStage.show();
    }

    private void importFile() {
        fileLabel.setText("");
        fileChooser.setTitle("Import File");
        List<File> files = fileChooser.showOpenMultipleDialog(primaryStage);

        if(files != null && !files.isEmpty()){
            StringBuilder filePaths = new StringBuilder();
            for(File file: files){
                String filePath = backend.importFile(file);
                if(!filePath.isEmpty()){
                    filePaths.append(filePath).append(", ");
                }
            }

            String allFilePaths = filePaths.toString();
            if(!allFilePaths.isEmpty()){
                allFilePaths = allFilePaths.substring(0, allFilePaths.length() - 2);
                fileLabel.setText(allFilePaths);
            }
        }
    }

    private void importImage(){
        fileLabel.setText("");
        fileChooser.setTitle("Import Image");
        File file = fileChooser.showOpenDialog(primaryStage);

        if(file != null){
            String filePath = backend.importImage(file);
            if(!filePath.isEmpty()){
                fileLabel.setText(filePath);
            }
        }

    }

    private void convertToDOCX(String filePaths) {
        String[] paths = filePaths.split(", ");
        List<String> docxFilePaths = new ArrayList<>();

        for (String path : paths) {
            String docxFilePath;
            if (isImageFile(path)) {
                docxFilePath = convertToDOCXImage(path);
            } else {
                docxFilePath = backend.convertToDOCX(path);
            }

            if (!docxFilePath.isEmpty()) {
                docxFilePaths.add(docxFilePath);
                convertedDocxFilePaths.add(docxFilePath);
            }
        }

        if (!docxFilePaths.isEmpty()) {
            downloadLink.setText("Download DOCX");
            downloadLink.setOnAction(e -> downloadDOCX(docxFilePaths));
            downloadLink.setVisible(true);
        } else {
            downloadLink.setVisible(false);
        }
    }

    private String convertToDOCXImage(String imagePath) {
        try {
            BufferedImage image = ImageIO.read(new File(imagePath));
            File docxFile = new File("converted.docx");

            XWPFDocument document = new XWPFDocument();
            ByteArrayOutputStream byteStream = new ByteArrayOutputStream();

            // Add the image to the document
            int width = image.getWidth();
            int height = image.getHeight();
            int imageType = BufferedImage.TYPE_INT_RGB;
            int resolution = 96; // Image resolution (in DPI)

            // Write the image data to the byte stream
            ImageIO.write(image, "jpg", byteStream);

            document.createParagraph().createRun().addPicture(new ByteArrayInputStream(byteStream.toByteArray()), XWPFDocument.PICTURE_TYPE_JPEG, "image.jpg", Units.toEMU(width), Units.toEMU(height));

            // Save the document to the byte array stream
            document.write(byteStream);
            document.close();

            // Save the byte array stream to the DOCX file
            try (FileOutputStream fos = new FileOutputStream(docxFile)) {
                byteStream.writeTo(fos);
            }

            return docxFile.getAbsolutePath();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }

        return "";
    }

    private boolean isImageFile(String filePath){
        String extension = filePath.substring(filePath.lastIndexOf(".") + 1);
        return extension.equalsIgnoreCase("png") || extension.equalsIgnoreCase("jpg") || extension.equalsIgnoreCase("gif");
    }
    private void downloadDOCX(List<String> docxFilePaths) {
        for(String docxFilePath: docxFilePaths){
            File file = new File(docxFilePath);
            if (file.exists()) {
                javafx.application.HostServices hostServices = getHostServices();
                hostServices.showDocument(file.getAbsolutePath());
            }
        }
    }

    private void clearConvertedFiles(){
        for(String docxFilePath: convertedDocxFilePaths){
            try{
                Path path = Paths.get(docxFilePath);
                Files.deleteIfExists(path);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        convertedDocxFilePaths.clear();
        downloadLink.setVisible(false);
    }



    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("CSV to DOCX Converter");

        Backend backend = new Backend();
        UI ui = new UI(primaryStage, backend);

        ui.show();
    }

}