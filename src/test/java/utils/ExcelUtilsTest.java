package utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtilsTest {
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
                XSSFSheet sheetSc = workbookSc.createSheet(student.name+"'s Schedule");
                Row rowSc= sheetSc.createRow(4);
                Cell cellSc = rowSc.createCell(2);
                cellSc.setCellValue(student.name);
                FileOutputStream fileout = new FileOutputStream("/Users/aliokursun/IdeaProjects/Excel/OutputFiles/"+student.name+"'s Schedule.xlsx");
                workbookSc.write(fileout);
                fileout.flush();
                fileout.close();


            }

        }


    }


}
