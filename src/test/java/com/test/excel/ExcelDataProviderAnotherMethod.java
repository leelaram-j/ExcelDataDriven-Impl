package com.test.excel;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

public class ExcelDataProviderAnotherMethod
{
    public static Object[][] getExcelData() throws Exception
    {
        DataFormatter formatter = new DataFormatter();
        FileInputStream fis = new FileInputStream("//Users//lee//Documents//ExcelDataDriven//DD_input.xlsx");
        XSSFWorkbook xs = new XSSFWorkbook(fis);
        XSSFSheet sheet = xs.getSheet("Sheet2");
        int rowCount = sheet.getPhysicalNumberOfRows();
        XSSFRow row = sheet.getRow(0);
        int colCount =row.getLastCellNum();
        Object data[][] = new Object[rowCount-1][colCount];
        for(int i=0;i<rowCount-1;i++)
        {
            row = sheet.getRow(i+1);
            for(int j=0;j<colCount;j++)
            {
                XSSFCell cell = row.getCell(j);
                data[i][j] = formatter.formatCellValue(cell);
            }
        }
        return data;
    }

}
