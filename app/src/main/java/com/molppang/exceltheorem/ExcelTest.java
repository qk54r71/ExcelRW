package com.molppang.exceltheorem;

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

/**
 * Created by park on 2016-07-20.
 */
public class ExcelTest {

    ArrayList<String[]> strings = new ArrayList<>();
    String fileName = "z";

    public static void main(String[] args) {
        try {
            new ExcelTest();

        } catch (Exception e) {

        }
    }

    public ExcelTest() throws IOException, Exception {
        copyExcelDataToDatabase();


     /*   File file = new File("C:\\Users\\park\\Desktop\\rr.xls");
        WritableWorkbook workbook = Workbook.createWorkbook(file);
        WritableSheet sheet = workbook.createSheet("Sheet", 0);

        int count = 0;
        for (int i = 0; i < 10; i++) {
            for (int j = 0; j < 10; j++, count++) {
                sheet.addCell(new Label(i, j, String.format("%d", count)));
            }
        }
        workbook.write();
        workbook.close();*/
    }

    private void copyExcelDataToDatabase() {

        Workbook workbook = null;
        Sheet sheet = null;

        System.out.println("시작");
        File readFile = new File("C:\\Users\\park\\Desktop\\jamo\\" + fileName + ".xls");


        try {
            workbook = workbook.getWorkbook(readFile);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }

        if (workbook != null) {
            sheet = workbook.getSheet(1);


            if (sheet != null) {

                int nMaxColumn = 4;
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
                    int j = 13;


                    for (int nColumn = nColumnStartIndex + 1; nColumn <= nColumnEndIndex; nColumn++) {
                        if (nRow % 6 == 0) {
                            break;
                        }
                        if (j == nColumn) {

                            ints[i++] = sheet.getCell(nColumn, nRow).getContents();
                            j++;
                            if (j == 18) {
                                break;
                            }
                        }

                    }
                    if (nRow % 6 == 0 || nRow ==1) {
                        strings.add(ints);
                    }
                }

                File file = new File("C:\\Users\\park\\Desktop\\complete\\" + fileName + ".xls");
                WritableWorkbook writableWorkbook = null;
                WritableSheet writableSheet = null;
                try {
                    writableWorkbook = workbook.createWorkbook(file);
                    writableSheet = writableWorkbook.createSheet("Sheet", 0);

                    for (int nRow = 0; nRow < strings.size(); nRow++) {
                        System.out.println("nRow : " + nRow + " strings.size() : " + strings.size());

                        writableSheet.addCell(new Label(0, nRow, strings.get(nRow)[14]));

                        for (int nColumn = 0; nColumn < strings.get(nRow).length; nColumn++) {
                            System.out.println("nRow : " + nRow + " nColumn : " + nColumn + " strings.size() : " + strings.size());
                            writableSheet.addCell(new Label(nColumn+1, nRow, strings.get(nRow)[nColumn]));
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
                } catch (NullPointerException e) {
                    System.out.println("Error : " + e.toString());
                } catch (Exception e) {
                    System.out.println("Error : " + e.toString());
                }
                workbook.close();
            }
        }
    }


}
