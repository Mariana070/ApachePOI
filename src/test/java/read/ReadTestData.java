package read;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadTestData {
    @Test
    public void getSpecificRows() throws IOException {
        //choose the rows that have value as Midwest
        File exelFile = new File("src/test/resources/Codefish-testData.xlsx");
        FileInputStream fis=new FileInputStream(exelFile);
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
        XSSFSheet page1 = workbook.getSheetAt(0);
        for (int i = page1.getFirstRowNum(); i <page1.getLastRowNum() ; i++) {
            XSSFRow tempRow = page1.getRow(i);
            if(i==0||tempRow.getCell(4).toString().equals("Midwest")){
                for (int j = tempRow.getFirstCellNum(); j <tempRow.getLastCellNum() ; j++) {
                }
            }

        }
    }
    @Test
    public void getSpecificRowsOptions2() throws IOException {
        //choose the values of Specific column(Business Type) Only,without hard-coding
        File exelFile = new File("src/test/resources/Codefish-testData.xlsx");
        FileInputStream fis =new FileInputStream(exelFile);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet page1 = workbook.getSheetAt(0);
        XSSFRow firstRow = page1.getRow(0);
        int index = -1;
        for (int i = firstRow.getFirstCellNum(); i < firstRow.getLastCellNum(); i++) {

            if(firstRow.getCell(i).toString().equals("BusinessType")){
                index =i;
            }
        }
        for (int i = page1.getFirstRowNum(); i <=page1.getLastRowNum() ; i++) {
            XSSFRow tempRow = page1.getRow(i);
            System.out.println(tempRow.getCell(index));

        }
    }
}
