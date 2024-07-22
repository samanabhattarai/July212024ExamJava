package FileInputOut;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

public class FileColumnInert1 {

    public static void main(String[] args) throws Exception {

        File file = new File("C:\\zorba_intellije\\MavenProject\\src\\main\\resources\\employee.xlsx");
        FileInputStream fis = new FileInputStream(file);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);

        Iterator ite = sheet.iterator();
        while (ite.hasNext()) {
            Row row = (Row) ite.next();

            Cell cell = null;

            if (row.getRowNum() == 0) {

                cell = row.createCell(row.getLastCellNum(), CellType.STRING);
                cell.setCellValue("Manager_id");
            } else {

                cell = row.createCell(row.getLastCellNum(), CellType.NUMERIC);

                switch (row.getRowNum()) {

                    case 1:
                        cell.setCellValue("null");
                        break;
                    case 2:
                        cell.setCellValue(1001);
                        break;
                    case 3:
                        cell.setCellValue(1004);
                        break;
                    case 4:
                        cell.setCellValue(1004);
                        break;
                    case 5:
                        cell.setCellValue(1001);
                        break;
                    case 6:
                        cell.setCellValue(1005);
                        break;
                    case 7:
                        cell.setCellValue(1001);
                        break;


                }


                if (row.getRowNum() == 0) {

                    cell = row.createCell(row.getLastCellNum(), CellType.STRING);
                    cell.setCellValue("emp_dept");
                } else {

                    cell = row.createCell(row.getLastCellNum(), CellType.STRING);

                    switch (row.getRowNum()) {


                        case 1:
                            cell.setCellValue("Finance");
                            break;
                        case 2:
                            cell.setCellValue("Finance");
                            break;
                        case 3:
                            cell.setCellValue("R&D");
                            break;
                        case 4:
                            cell.setCellValue("R&D");
                            break;
                        case 5:
                            cell.setCellValue("Finance");
                            break;
                        case 6:
                            cell.setCellValue("Finance");
                            break;
                        case 7:
                            cell.setCellValue("Finance");
                            break;


                    }

                }
                FileOutputStream fos = new FileOutputStream(file);
                xssfWorkbook.write(fos);
                fos.close();
                System.out.println("Sucessfully write back to excel file");


            }


        }
    }
}
