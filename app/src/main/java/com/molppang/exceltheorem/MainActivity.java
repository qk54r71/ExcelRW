package com.molppang.exceltheorem;

import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class MainActivity extends AppCompatActivity {
    private int[][] values = new int[17][];
    private ArrayList<String[]> strings = new ArrayList<String[]>();

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        copyExcelDataToDatabase();

    }

    private void copyExcelDataToDatabase() {

        Workbook workbook = null;
        Sheet sheet = null;


        InputStream is = null;
        try {
            is = getBaseContext().getResources().getAssets().open("r.xls");
        } catch (IOException e) {
            e.printStackTrace();
        }


        try {
            workbook = workbook.getWorkbook(is);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }

        if (workbook != null) {
            sheet = workbook.getSheet(1);


            if (sheet != null) {

                int nMaxColumn = 2;
                int nRowStartIndex = 1;
                int nRowEndIndex = sheet.getColumn(nMaxColumn - 1).length - 1;
                int nColumnStartIndex = 0;
                int nColumnEndIndex = sheet.getRow(2).length - 1;


                int i = 0;
                String[] ints = new String[25];
                for (int nRow = nRowStartIndex; nRow <= nRowEndIndex; nRow++) {

                    if (i == 25) {
                        i = 0;
                        ints = new String[25];
                    }
                    int j = 7;


                    for (int nColumn = nColumnStartIndex + 1; nColumn <= nColumnEndIndex; nColumn++) {
                        if (nRow % 6 == 0) {
                            break;
                        }
                        if (j == nColumn) {
                            Log.i("String", "nColumn : " + nColumn + " nRow : " + nRow + " data : " + sheet.getCell(nColumn, nRow).getContents());

                            ints[i++] = sheet.getCell(nColumn, nRow).getContents();
                            Log.i("String", "ints " + (i - 1) + " : " + ints[i - 1]);
                            j++;
                            if (j == 12) {
                                break;
                            }
                        }

                    }
                    strings.add(ints);
                }

                for (String s : strings.get(0)) {
                    Log.i("StringData", "" + s);
                }

                File file = new File("file:///android_asset/rr.xls");
                WritableWorkbook writableWorkbook = null;
                WritableSheet writableSheet = null;
                try {
                    writableWorkbook = workbook.createWorkbook(file);
                    writableSheet = writableWorkbook.createSheet("Sheet", 0);

                    for (int nRow = 0; nRow <= strings.size(); nRow++) {

                        for (int nColumn = 0; nColumn <= strings.get(nRow).length; nColumn++) {
                            writableSheet.addCell(new Label(nRow, nColumn, strings.get(nRow)[nColumn]));
                        }

                    }
                    writableWorkbook.write();
                    writableWorkbook.close();

                } catch (IOException e) {
                    e.printStackTrace();
                } catch (RowsExceededException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }


            } else {
            }
        } else {
            if (workbook != null) {
                workbook.close();
            }


        }


    }
}