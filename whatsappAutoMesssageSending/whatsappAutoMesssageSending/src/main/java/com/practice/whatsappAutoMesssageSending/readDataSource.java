package com.practice.whatsappAutoMesssageSending;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class readDataSource {


    public String ReadCellData(int vRow, int vColumn){
        String value = null;
        Workbook wb = null;
        try {
            File currentDirFile = new File(".//students.xlsx");
            FileInputStream fis = new FileInputStream(currentDirFile);
            wb = new XSSFWorkbook(fis);
        } catch (IOException e1) {
            e1.printStackTrace();
        }
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(vRow);
        Cell cell = row.getCell(vColumn);
        value = cell.getStringCellValue();
        return value;
    }

    public static void main(String args[]) throws IOException{
        readDataSource rd = new readDataSource();
        for (int i = 1; i <5; i++) {
            String vOutput = rd.ReadCellData(i, 1);
            System.out.println(vOutput);
        }
    }
}
