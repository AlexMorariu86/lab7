package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class Main {
    public static void main(String[] args)
    {

        try {

            FileInputStream file = new FileInputStream(new File("C:\\Users\\senon\\Desktop\\Book1.xlsx"));

            XSSFWorkbook workbook = new XSSFWorkbook(file);

            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {

                        case NUMERIC:
                            System.out.print(
                                    cell.getNumericCellValue()
                                            + "t");
                            break;

                        case STRING:
                            System.out.print(
                                    cell.getStringCellValue()
                                            + "t");
                            break;
                    }
                }

                System.out.println("");
            }

            file.close();
        }

        catch (Exception e) {

            e.printStackTrace();
        }
    }
}