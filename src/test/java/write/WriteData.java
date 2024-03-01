package write;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;

public class WriteData {
    @Test
    public void writeDataTest() throws IOException {
        File file= new File("src/test/resources/b16Test.xlsx");
        FileInputStream fis =new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet page1 = workbook.getSheetAt(0);
        //create a new row after last one
        XSSFRow newRow= page1.createRow(page1.getLastRowNum()+1);
        XSSFCell newCell= newRow.createCell(0);
        newCell.setCellType(CellType.NUMERIC);
        newCell.setCellValue(6);
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos); // write value '6' into this file

    }
}
