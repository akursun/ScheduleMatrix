package utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtilsTest {
    static Cell cell;
    static CellStyle cellStyle;
    public static void main(String[] args) throws IOException {
        String excelPath = "./data/StudentSchedule.xlsx";
        String sheetName = "2021-2022 HS Student Schedule";
        ExcelUtils excel = new ExcelUtils(excelPath, sheetName);
        excel.getRowCount();
        excel.getCellData(1, 0);
        excel.getCellData(1, 1);
        excel.getCellData(1, 2);
        for (int i = 0; i < excel.workbook.getSheet(sheetName).getPhysicalNumberOfRows()-1; i++) {
            String firstCellValue = excel.getCellData(i, 0);
            if (firstCellValue.equals("Student:")) {
                Student student = new Student();
                student.name = excel.getCellData(i, 1);
                student.number = excel.getCellData(i, 3);
                student.grade = excel.getCellData(i, 5);
                student.lockerCombination = excel.getCellData(i, 11);
                System.out.println(student.name);
                System.out.println(student.number);
                System.out.println(student.grade);
                System.out.println(student.lockerCombination);
                System.out.println("----------");
                //TODO Get student's classes.

                XSSFWorkbook workbookSc = new XSSFWorkbook();
                XSSFSheet sheetSc = workbookSc.createSheet("Schedule");
                Row row = sheetSc.createRow(0);

                cell = row.createCell(0);
                cell.setCellValue("Student:");
                cell = row.createCell(1);
                cell.setCellValue(student.name);

                row = sheetSc.createRow(1);
                cell = row.createCell(0);
                cell.setCellValue("ID:");
                cell = row.createCell(1);
                cell.setCellValue(Integer.parseInt(student.number));

                row = sheetSc.createRow(2);
                cell = row.createCell(0);
                cell.setCellValue("Grade:");
                cell = row.createCell(1);
                cell.setCellValue(Integer.parseInt(student.grade));

                row = sheetSc.createRow(3);
                cell = row.createCell(0);
                cell.setCellValue("Locker:");
                cell = row.createCell(1);
                cell.setCellValue(student.lockerCombination);





                FileOutputStream fileout = new FileOutputStream("/Users/aliokursun/IdeaProjects/Excel/OutputFiles/"+student.name+"'s Schedule.xlsx");
                workbookSc.write(fileout);
                fileout.flush();
                fileout.close();


            }

        }


    }



}
