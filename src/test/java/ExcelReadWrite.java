import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadWrite {

    public static void main(String[] args) {
        String[] headers = {"Name", "Age", "Email"};
        String[][] data = {
                {"John Doe", "30", "john@test.com"},
                {"Jane Doe", "28", "jane@test.com"},
                {"Bob Smith", "35", "bob@test.com"},
                {"Swapnil", "37", "swapnil@example.com"}
        };

        // Write data to Excel
        writeExcel(headers, data);

        // Read data from Excel and print to console
        readExcelAndPrint();
    }

    public static void writeExcel(String[] headers, String[][] data) {
        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }
            for (int i = 0; i < data.length; i++) {
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < data[i].length; j++) {
                    row.createCell(j).setCellValue(data[i][j]);
                }
            }
            FileOutputStream fileOut = new FileOutputStream("output.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            System.out.println("Excel file has been created successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void readExcelAndPrint() {
        try {
            FileInputStream fis = new FileInputStream("output.xlsx");
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(cell.getStringCellValue() + "\t");
                }
                System.out.println();
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
