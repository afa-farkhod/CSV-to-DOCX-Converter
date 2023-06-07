# CSV-to-DOCX-Converter
Desktop application that allows users to import files (*.csv files and *.png, *.jpg), convert them to DOCX format, and download the converted files.

- Main architecture, GUI allows users to import files (including CSV files and images), convert them to DOCX format, and download the converted files.

![image](https://github.com/af4092/CSV-to-DOCX-Converter/assets/24220136/7aefc5b9-f253-4726-aff3-ac9022abedb5)

- JavaFX framework based desktop application main view:
-  UI elements and Backend functions are separated.
- `Import File` button imports single and multiple `*.csv` files (successfully tested with 7 *.csv files) at the same time. 
- `Convert to DOCX` button enables to convert and download multiple `*.docx` files at the same time. 
- `Import Image` button enables to import image and move it to Word.
- `CLEAR` button clears the local directory from created files
- `Exit` button which stops and exits the GUI application

![image](https://github.com/af4092/CSV-to-DOCX-Converter/assets/24220136/a86bd5d0-0a3d-4383-8b7a-546fed7fe6da)

- Demo video shows the simulation of importing several files and clear them from the local directory, lastly shows the importing image and moving it to word:

https://github.com/af4092/CSV-to-DOCX-Converter/assets/24220136/e090273e-a3df-4e51-9f3a-8c9ec7db9d80

