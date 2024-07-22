package FileInputOut;

import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

public class ExcelFileRead {
    public static void main(String[] args) throws IOException {

        Scanner scanner = new Scanner(System.in);
        System.out.print("enter employee_id: ");
        int empid = scanner.nextInt();



     /*   int numrow = sheet.getLastRowNum();


        Iterator<Row> rowit = sheet.iterator();
        while (rowit.hasNext()) {
            Row row = rowit.next();
            int cellnum = row.getLastCellNum();
            Iterator<Cell> cellit = row.iterator();
            while (cellit.hasNext()) {
                Cell cell = cellit.next();
                int cellpos = cell.getColumnIndex();
                System.out.print("cell at index: "+ cellpos + " ");
                String cellType = cell.getCellType().name();
                switch (cellType) {
                    case "STRING" :
                        System.out.print(cell.getStringCellValue());
                        break;
                    case "NUMERIC":
                        System.out.print(cell.getNumericCellValue());
                        break;
                    case "BOOLEAN":
                        System.out.print(cell.getBooleanCellValue());
                        break;
                }
                System.out.println();
            }
        }*/


        try {
            FileInputStream file = new FileInputStream("C:\\zorba_intellije\\MavenProject\\src\\main\\resources\\employee.xlsx");

            XSSFWorkbook workbook = new XSSFWorkbook(file);

            Sheet sheet = workbook.getSheetAt(0);


            // Read the data from the Excel file
            for (Row dataRow : sheet) {

                int rowEmployeeId = (int) dataRow.getCell(0).getNumericCellValue();

                if (rowEmployeeId == empid) {
                    System.out.println("EmpID: " + rowEmployeeId);

                    System.out.println("EmpName: " + dataRow.getCell(1).getStringCellValue());

                    System.out.println("EmpSalay: " + dataRow.getCell(2).getNumericCellValue());

                    System.out.println("EmpMobile: " + dataRow.getCell(3).getStringCellValue());

                    System.out.println("EmpCity: " + dataRow.getCell(4).getStringCellValue());

                    System.out.println("ManagerID: " + dataRow.getCell(5).getNumericCellValue());

                    System.out.println("EmpDept: " + dataRow.getCell(6).getStringCellValue());

                    System.out.println("EmpShare: " + dataRow.getCell(7).getNumericCellValue());

                    break;
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
            System.out.println(e.getMessage());
            System.out.println("File not found or file has some exceptions");
        }
        }
    }


