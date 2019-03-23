package com.test.excel;

import java.io.IOException;
import java.util.ArrayList;

public class DDTesting
{
    public static void main(String[] args)throws IOException
    {
        String fileName="//Users//lee//Documents//DD_input.xlsx";
        String sheetName="Sheet1";
        String headerName="TestCases";
        String testCase="Purchase";

        ArrayList<String> result = DataDriven.getData(fileName,headerName,sheetName,testCase);
        System.out.println(result.get(1));
        System.out.println(result.get(2));
        System.out.println(result.get(3));
    }
}
