package utils;

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
            }
        }


    }
}
