# CSV-to-DOCX-Converter
Desktop application that allows users to import files (`*.csv` files and `*.png`, `*.jpg`), convert them to `DOCX` format, and download the converted files.

- [Filestart.com](https://filestar.com/skills/csv/convert-csv-to-docx) - Filestar is a desktop application that runs on both Windows and OSX, ensuring that you can use it no matter what operating system you prefer. The software is perfect for professionals who need to convert data into professional-looking documents, such as reports, invoices, and proposals.
- Main architecture, GUI allows users to import files (including CSV files and images), convert them to DOCX format, and download the converted files.

<p align="center">
  <img src="https://github.com/af4092/CSV-to-DOCX-Converter/assets/24220136/7aefc5b9-f253-4726-aff3-ac9022abedb5" alt="Image">
</p>

- JavaFX framework based desktop application main view:
  -  UI elements and Backend functions are separated.
  - `Import File` button imports single and multiple `*.csv` files (successfully tested with 7 *.csv files) at the same time. 
  - `Convert to DOCX` button enables to convert and download multiple `*.docx` files at the same time. 
  - `Import Image` button enables to import image and move it to Word.
  - `CLEAR` button clears the local directory from created files
  - `Exit` button which stops and exits the GUI application

<p align="center">
  <img src="https://github.com/af4092/CSV-to-DOCX-Converter/assets/24220136/a86bd5d0-0a3d-4383-8b7a-546fed7fe6da" alt="Image">
</p>

- The code is organized in the following package `com.example.csvtodocxconverter`;
- The main class is called UI, representing the user interface of the application.
- The UI class extends the `javafx.application.Application` class, which is the base class for JavaFX applications.
- The class defines various instance variables that hold references to UI components such as the stage, file chooser, labels, buttons, and a backend object.
- The class contains constructors that initialize the instance variables and set up event handlers for the buttons.
- The `show()` method sets up the user interface layout using a vertical box (VBox) container and displays it in the application window.
- The class provides methods for importing files, both CSV files (importFile()) and images (importImage()), and displaying the selected file paths in the UI.
- The `convertToDOCX()` method converts the selected files to the DOCX format, either by utilizing the backend object's convertToDOCX() method or by converting image files to DOCX using Apache POI.
- The `downloadDOCX()` method allows users to download the converted DOCX files by opening them with the default system application.
- The `clearConvertedFiles()` method deletes the converted DOCX files from the system and clears the list of converted file paths.
- The `start()` method is the entry point of the JavaFX application and initializes the application window, backend object, and user interface.
- The `main()` method launches the JavaFX application.

- Demo video shows the simulation of importing several files and clear them from the local directory, lastly shows the importing image and moving it to word:

https://github.com/af4092/CSV-to-DOCX-Converter/assets/24220136/e090273e-a3df-4e51-9f3a-8c9ec7db9d80
