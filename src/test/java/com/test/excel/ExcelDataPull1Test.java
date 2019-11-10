package com.test.excel;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataPull1Test
{
    @Test(dataProvider = "driveTest")
    public void excelTestData(String greeting, String communication, String id)
    {
        System.out.println(greeting+" "+communication+" "+id);

    }

    @DataProvider(name="driveTest")
    public Object[][] getData() throws Exception
    {
        Object[][] obj = ExcelDataProviderAnotherMethod.getExcelData();
        return obj;
    }
}
