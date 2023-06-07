module com.example.csvtodocxconverter {
    requires javafx.controls;
    requires javafx.fxml;
    requires poi.ooxml;
    requires poi;
    requires java.desktop;


    opens com.example.csvtodocxconverter to javafx.fxml;
    exports com.example.csvtodocxconverter;
}