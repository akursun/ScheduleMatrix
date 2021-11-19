package utils;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class ExcelUtils {

    static XSSFWorkbook workbook;
    static Sheet sheet;


    public ExcelUtils(String excelPath, String sheetName) {
        try {
            workbook = new XSSFWorkbook(excelPath);
            sheet = workbook.getSheet(sheetName);
        } catch (Exception exp) {
            System.out.println(exp.getCause());
            System.out.println(exp.getMessage());
            exp.printStackTrace();
        }
    }

    public static String getCellData(int rowNum, int colNum) throws IOException {
        DataFormatter dataFormatter = new DataFormatter();
        Object value = dataFormatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
        return value.toString();
    }

    public static void getRowCount() {
        int rowCount = sheet.getPhysicalNumberOfRows();
        System.out.println("Number of Rows = " + rowCount);
    }







}
