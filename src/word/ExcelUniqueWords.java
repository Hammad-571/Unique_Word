package word;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

public class ExcelUniqueWords {

    public static void main(String[] args) {
        try {
            FileInputStream excelFile = new FileInputStream(new File("input.xlsx")); // Replace with your input Excel file path
            Workbook workbook = new XSSFWorkbook(excelFile);

            Set<String> uniqueWords = new HashSet<>();

            // Assuming the data is in the first sheet (index 0)
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through each row in the sheet
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Iterate through each cell in the row
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        String[] words = cellValue.split("\\s+"); // Split the cell value into words

                        for (String word : words) {
                            uniqueWords.add(word);
                        }
                    }
                }
            }

            // Create a new workbook to write the unique words
            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("UniqueWords");

            // Write unique words to the new sheet
            int rownum = 0;
            for (String word : uniqueWords) {
                Row newRow = newSheet.createRow(rownum++);
                Cell cell = newRow.createCell(0);
                cell.setCellValue(word);
            }

            // Save the new workbook to a file
            FileOutputStream outputStream = new FileOutputStream(new File("output.xlsx")); // Replace with your output file path
            newWorkbook.write(outputStream);
            outputStream.close();

            System.out.println("Unique words have been written to the output file.");

            excelFile.close();
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
