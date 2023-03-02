package com.example.saveinexceluseingjava;


import static androidx.constraintlayout.helper.widget.MotionEffect.TAG;

import android.annotation.SuppressLint;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class MainActivity extends AppCompatActivity {

    private TextView btSave;

    @SuppressLint("MissingInflatedId")
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        btSave = findViewById(R.id.savedata);
        btSave.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                savefile();
            }
        });
    }

    public void savefile() {

        String[] headers = {"Sura Name", "Translation", "Aya", "Sura No", "Sura Name (English)", "Text"};
        Object[][] data = {{"ٱلْبَقَرَة", "Who believe in the Unknown and fulfil their devotional obligations, and spend in charity of what We have given them; ", "3", "2", "Al-Baqarah", "الَّذِينَ يُؤْمِنُونَ بِالْغَيْبِ وَيُقِيمُونَ الصَّلَاةَ وَمِمَّا رَزَقْنَاهُمْ يُنْفِقُونَ"}};

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Report");
        Row headerRow = sheet.createRow(0);

        //for header
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }
        //for set data into cell in excel
        for (int i = 0; i < data.length; i++) {
            Row dataRow = sheet.createRow(i + 1);
            for (int j = 0; j < data[i].length; j++) {
                Cell cell = dataRow.createCell(j);
                cell.setCellValue(data[i][j].toString());
            }
        }

        //set a internal path
        String yourFilePath = this.getFilesDir() + "/" + "hello.xls";
        File yourFile = new File(yourFilePath);

        try {
            FileOutputStream fileOut = new FileOutputStream(yourFile); //Opening the file
            workbook.write(fileOut); //Writing all your row column inside the file
            fileOut.close(); //closing the file and done
            shareapp(yourFile);  //share a file into other app
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public void shareapp(File file) {

        Log.e(TAG, "I shouldn't be here" + file);
        Log.e(TAG, "file path" + file.toString());

        final Intent intent = new Intent(android.content.Intent.ACTION_SEND);
        intent.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
        Uri photoURI = FileProvider.getUriForFile(getApplicationContext(), BuildConfig.APPLICATION_ID + ".provider", file);
        intent.putExtra(Intent.EXTRA_STREAM, photoURI);
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        intent.setType("image/png");
        startActivity(Intent.createChooser(intent, "Share image via"));
    }


    //exm- 1
    /*public void savefile() {

        String[] headers = {"Date","Expense Category", "Amount", "Notes","Account Number", "Paid By"};
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Users"); //Creating a sheet

        // Create a new row for headers
        Row headerRow = sheet.createRow(0);

        // Set the headers
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        Row dataRow = sheet.createRow(1);
        dataRow.createCell(0).setCellValue("2023-02-28");
        dataRow.createCell(1).setCellValue("Food");
        dataRow.createCell(2).setCellValue(10.50);
        dataRow.createCell(3).setCellValue("Lunch with colleagues");
        dataRow.createCell(4).setCellValue("123456");
        dataRow.createCell(5).setCellValue("John Smith");


        String fileName = "FileName.xlsx"; //Name of the file
        //String extStorageDirectory = Environment.getExternalStorageDirectory().toString();

        String yourFilePath = this.getFilesDir() + "/" + "hello.xls";
        File yourFile = new File( yourFilePath );

        try {
            FileOutputStream fileOut = new FileOutputStream(yourFile); //Opening the file
            workbook.write(fileOut); //Writing all your row column inside the file
            fileOut.close(); //closing the file and done
            shareapp(yourFile);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/

    /*public void shareapp(File file){

         Log.e(TAG,"I shouldn't be here"+file);
         Log.e(TAG,"file path"+file.toString());

        final Intent intent = new Intent(android.content.Intent.ACTION_SEND);
        intent.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
        Uri photoURI = FileProvider.getUriForFile(getApplicationContext(), BuildConfig.APPLICATION_ID +".provider", file);
        intent.putExtra(Intent.EXTRA_STREAM, photoURI);
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        intent.setType("image/png");
        startActivity(Intent.createChooser(intent, "Share image via"));
    }*/
}