package read;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadData {
    @Test
    public void readDataTest() throws IOException {

        File exelFile = new File("src/test/resources/b16Test.xlsx");
        FileInputStream fis=new FileInputStream(exelFile);
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
        XSSFSheet page1 = workbook.getSheetAt(0);
        XSSFRow row1 = page1.getRow(0);
        XSSFCell cell1 = row1.getCell(0);
        for (int i = row1.getFirstCellNum(); i < row1.getLastCellNum(); i++) {
            XSSFCell tempCell = row1.getCell(i);
        }
    }
    @Test
    public void getAllDataTest() throws IOException {
        File exelFile = new File("src/test/resources/b16Test.xlsx");
        FileInputStream fis = new FileInputStream(exelFile);
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
        XSSFSheet page1=workbook.getSheet("Sheet1");
        //iterate over each row in Excel file
        for (int i = page1.getFirstRowNum(); i <=page1.getLastRowNum() ; i++) {
            XSSFRow tempRow=page1.getRow(i);
            //iterate over each cell in tempRow
            for (int j = tempRow.getFirstCellNum(); j <tempRow.getLastCellNum() ; j++) {
            }
        }
    }


}
