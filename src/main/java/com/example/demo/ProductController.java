package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.util.ArrayList;

@RestController
public class ProductController {

    @PostMapping("/product")
    public void printData(@RequestBody Product product) throws IOException {
       //converter.my();
//        ArrayList <String> names = new ArrayList<>();
//        names.add(product.getProduct());
//        names.add(product.getCountry());
//        names.add(String.valueOf(product.getBrest()));



      //  File xlsxFile = new File("C:/Users/evgen/OneDrive/Рабочий стол/123.xlsx");
      //  FileInputStream inputStream = new FileInputStream(xlsxFile);
        Workbook workbook = new XSSFWorkbook();
     //   FileOutputStream os = new FileOutputStream(xlsxFile);

        Sheet sheet1 = workbook.createSheet(String.valueOf(1));
        Cell cell_product;
        Cell cell_country;
        Cell cell_brest;
        Cell cell_vitebsk;
        Cell cell_minsk;
        Cell cell_mogilev;

       int i = sheet1.getLastRowNum()+1;

            Row row = sheet1.createRow(i);
            cell_product = sheet1.getRow(i).getCell(i);
            cell_product = row.createCell(i);
            cell_product.setCellValue(product.getProduct());

            cell_country = sheet1.getRow(i).getCell(i+1);
            cell_country = row.createCell(i+1);
            cell_country.setCellValue(product.getCountry());

            cell_brest = sheet1.getRow(i).getCell(i+2);
            cell_brest = row.createCell(i+2);
            cell_brest.setCellValue(product.getBrest());

            cell_vitebsk = sheet1.getRow(i).getCell(i+3);
            cell_vitebsk = row.createCell(i+3);
            cell_vitebsk.setCellValue(product.getVitebsk());

            cell_minsk = sheet1.getRow(i).getCell(i+4);
            cell_minsk = row.createCell(i+4);
            cell_minsk.setCellValue(product.getMinsk());

            cell_mogilev = sheet1.getRow(i).getCell(i+5);
            cell_mogilev = row.createCell(i+5);
            cell_mogilev.setCellValue(product.getMogilev());



        FileOutputStream fileOut = new FileOutputStream("C:/Users/evgen/OneDrive/Рабочий стол/1.xlsx");
        workbook.write(fileOut);
        fileOut.close();

       // workbook.write(os);
       // workbook.close();

        System.out.println("Excel file has been updated successfully.");
        System.out.println("Printing the product data:"+product.getCountry()  + product.getProduct()+ product.getBrest());




        }
}