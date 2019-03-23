package com.test.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class DataDriven
{
    //Identify the columns in which the test case name is there. scanning entire sheet
    //Once the column with entire test cases is identified, then we need to go to the column with the test case name
    //After reaching the column we need to get the list of data that are available for this particular row/test case
    //feed the same to test case
    public static ArrayList<String> getData(String fileName, String headerName, String sheetName, String testCaseName) throws IOException
    {
        ArrayList<java.lang.String> a = new ArrayList<java.lang.String>();

        // get the file into fileinputstream object
        FileInputStream fis = new FileInputStream(fileName);
        // pass the fis object into xssfworkbook. this will open the excel file
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        //get the sheet count so that you can Iterate through it
        int sheetCount = workbook.getNumberOfSheets();
        for (int i = 0; i < sheetCount; i++)
        {
            //Identify the sheet by the sheet name
            if (workbook.getSheetName(i).equalsIgnoreCase(sheetName))
            {
                //pass that sheet to a XSSFSheet object.
                XSSFSheet sheet = workbook.getSheetAt(i);
                Iterator<Row> rows = sheet.iterator();
                //this will send the control to rows 1,1;
                Row firstRow = rows.next();
                Iterator<Cell> cell = firstRow.cellIterator();
                int k = 0;
                int column = 0;
                while (cell.hasNext())
                {
                    Cell value = cell.next();
                    if (value.getStringCellValue().equalsIgnoreCase(headerName))
                    {
                        column = k;
                    }
                    k++;
                }
                //System.out.println(column);
                while (rows.hasNext())
                {
                    Row r = rows.next();
                    if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName))
                    {
                        //after u grab the specific row, get all the data in that row
                        Iterator<Cell> tv = r.cellIterator();
                        while (tv.hasNext())
                        {
                            Cell checkData = tv.next();
                            if(checkData.getCellTypeEnum()==CellType.STRING)
                            {
                                a.add(checkData.getStringCellValue());
                            }
                            else
                                {
                                a.add(NumberToTextConverter.toText(checkData.getNumericCellValue()));
                            }
                        }
                    }
                }
            }

        }
        return a;
    }

}
