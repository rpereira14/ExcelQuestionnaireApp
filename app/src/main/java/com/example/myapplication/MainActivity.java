package com.example.myapplication;

import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    Button excel;
    Button update;
    Button read;
   static TextView text;

    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        excel = (Button)findViewById(R.id.Excel);
        excel.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Log.d("TAG",""+ saveExcelFile("MyExcel.xls"));
                saveExcelFile("MyExcel.xls");
            }
        });
        update = (Button)findViewById(R.id.UpdateExcel);
        update.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                updateExcelFile();
            }
        });
        text = (TextView) findViewById(R.id.TextView);
        read = (Button) findViewById(R.id.ReadExcel);
        read.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                readExcelFile();

            }
        });

    }

    private static boolean saveExcelFile(String fileName) {
        String path;
        File dir;
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.e("Failed", "Storage not available or read only");
            return false;
        }
        boolean success = false;

        //New Workbook
        Workbook wb = new HSSFWorkbook();

        Cell c = null;

        //Cell style for header row

        /*CellStyle cs = wb.createCellStyle();
        cs.setFillForegroundColor(HSSFColor.LIME.index);
        cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
        */

        //New Sheet
        Sheet sheet1 = null;
        sheet1 = wb.createSheet("myOrder");

        // Generate column headings
        Row row = null;

        row = sheet1.createRow(0);

        c = row.createCell(0);
        c.setCellValue("Item Number");
        // c.setCellStyle(cs);

        c = row.createCell(1);
        c.setCellValue("Quantity");
        //  c.setCellStyle(cs);

        c = row.createCell(2);
        c.setCellValue("Price");
        // c.setCellStyle(cs);

        sheet1.setColumnWidth(0, (15 * 500));
        sheet1.setColumnWidth(1, (15 * 500));
        sheet1.setColumnWidth(2, (15 * 500));

        int val = 0;
        int k = 1;
        for(int i=1;i<12;i++){
            row = sheet1.createRow(k);
            for(int j=0;j<3;j++){
                c = row.createCell(j);
                c.setCellValue(val);
                //c.setCellStyle(cellStyle);
                val++;
            }
            sheet1.setColumnWidth(i, (15 * 500));
            k++;
        }

        path = Environment.getExternalStorageDirectory().getAbsolutePath()+"/EXCEL/";
        dir = new File(path);
        Log.d("TAG","" + path);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        File file = new File(dir, fileName);
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            Log.w("FileUtils", "Writing file" + file);
            success = true;
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
            }
        }
        return success;
    }

    public static void updateExcelFile(){
        String excelFilePath = Environment.getExternalStorageDirectory().getAbsolutePath()+"/EXCEL/" + "MyExcel.xls";
        File dir;
        try {
            FileInputStream inputStream = new FileInputStream(
                    new File(new File(Environment.getExternalStorageDirectory().getAbsolutePath()+"/EXCEL/"),"MyExcel.xls"));
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheet("myOrder");

            int rowCount = sheet.getLastRowNum();
            Log.d("TAG",""+rowCount);
            int val = 1000;
            for(int i=1;i<5;i++){
                Row row = sheet.createRow(++rowCount);
                Cell cell = null;
                for(int j=0;j<3;j++){
                    cell = row.createCell(j);
                    cell.setCellValue(val);
                    //c.setCellStyle(cellStyle);
                    val +=10;
                }
                sheet.setColumnWidth(i, (15 * 500));
            }
            inputStream.close();
            dir = new File(Environment.getExternalStorageDirectory().getAbsolutePath()+"/EXCEL/");
            File file = new File(dir,"MyExcel.xls");
            FileOutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }
    public static void readExcelFile(){
        String excelFilePath = Environment.getExternalStorageDirectory().getAbsolutePath()+"/EXCEL/" + "MyExcel.xls";
        File dir;
        try {
            FileInputStream inputStream = new FileInputStream(
                    new File(new File(Environment.getExternalStorageDirectory().getAbsolutePath()+"/EXCEL/"),"MyExcel.xls"));
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheet("myOrder");
            String texto = null;
            int rowCount = 0;
            int val = 1000;
            for(int i=1;i<5;i++){
                Row row = sheet.getRow(1);
                Cell cell = null;
                for(int j=1;j<3;j++){
                    cell = row.getCell(j);
                    String stringCell = String.valueOf(cell.getNumericCellValue());
                    texto += " " +stringCell;
                    cell.setCellValue(val);
                    //c.setCellStyle(cellStyle);
                    val +=10;
                }
                sheet.setColumnWidth(i, (15 * 500));
            }
            text.setText(texto);
            inputStream.close();


            workbook.close();


        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }

    public static boolean isExternalStorageReadOnly() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) {
            return true;
        }
        return false;
    }

    public static boolean isExternalStorageAvailable() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) {
            return true;
        }
        return false;
    }
}