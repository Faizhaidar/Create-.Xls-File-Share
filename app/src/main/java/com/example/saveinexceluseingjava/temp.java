package com.example.saveinexceluseingjava;

import android.os.Environment;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class temp

{
    public void ss(){

        String[] headers = {"Date","Expense Category", "Amount", "Notes","Account Number", "Paid By"};

        // Create a new workbook
        Workbook workbook = new HSSFWorkbook();

        // Create a new worksheet
        Sheet sheet = workbook.createSheet("Expenses");

        // Create a new row for headers
        Row headerRow = sheet.createRow(0);

        // Set the headers
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // Add data to the worksheet
        // For example, add an expense record
        Row dataRow = sheet.createRow(1);
        dataRow.createCell(0).setCellValue("2023-02-28");
        dataRow.createCell(1).setCellValue("Food");
        dataRow.createCell(2).setCellValue(10.50);
        dataRow.createCell(3).setCellValue("Lunch with colleagues");
        dataRow.createCell(4).setCellValue("123456");
        dataRow.createCell(5).setCellValue("John Smith");


        String fileName = "FileName.xlsx";
        String extStorageDirectory = Environment.getExternalStorageDirectory().toString();
        File folder = new File(extStorageDirectory, "FolderName");
        folder.mkdir();
        File file = new File(folder, fileName);
        try {
            file.createNewFile();
        } catch (IOException e1) {
            e1.printStackTrace();
        }


        try {
            FileOutputStream fileOut = new FileOutputStream(file);
            workbook.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

         /*   // Save the workbook to a file
            String fileName = "expenses.xls";
            FileOutputStream outputStream = null;
            try {
                FileOutputStream outputStream = new FileOutputStream(fileName);
                workbook.write(outputStream);
            }catch (IOException e){
                e.printStackTrace();
            }finally {
                if (outputStream != null) {
                    try {
                        outputStream.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
            *//*   workbook.close();
            outputStream.close();*//*

            System.out.println("Data has been saved to " + fileName);*/
    }
}
