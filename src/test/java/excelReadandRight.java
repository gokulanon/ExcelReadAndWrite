import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelReadandRight {

    public static void main(String[] args) {
        String[] headers = {"Name", "Age", "Email"};
        String[][] data = {
            {"John Doe", "30", "john@test.com"},
            {"Jane Doe", "28", "jane@test.com"},
            {"Bob Smith", "35", "bob@test.com"},
            {"Swapnil", "37", "swapnil@example.com"}
        };

        try {
            // Create a new Excel workbook
            Workbook workbook = new XSSFWorkbook();

            // Create a new sheet with the name "Sheet1"
            Sheet sheet = workbook.createSheet("Sheet1");

            // Write column headers
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Write data rows
            for (int i = 0; i < data.length; i++) {
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < data[i].length; j++) {
                    row.createCell(j).setCellValue(data[i][j]);
                }
            }

            // Write the workbook to a file
            FileOutputStream fileOut = new FileOutputStream("output.xlsx");
            workbook.write(fileOut);
            fileOut.close();

            // Close the workbook
            workbook.close();
            System.out.println("Excel file has been created successfully!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
