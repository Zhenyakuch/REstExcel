package com.example.demo;

import com.example.demo.model.TotalMass;
import com.example.demo.model.dto.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.time.LocalDate;

@RestController
@Slf4j
public class ProductController {

    @PostMapping("/import-export")
    public String printData(@RequestBody CountryReport countryRequest) throws Exception {

        log.debug("CountryReport " + countryRequest);
        TotalMass totalMass = new TotalMass(countryRequest);
        log.debug("TotalMass " + totalMass);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook("src/main/resources/Blanc.xlsx");
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStyle.setWrapText(true);

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        //font.setItalic(true);
        cellStyle.setFont(font);


        String fromTitle = sheet.getRow(0).getCell(0).toString();
        String fromTitle2 = sheet.getRow(1).getCell(1).toString();

        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

        for (int i = 2; i <= 17; i++) {
            String date2 = sheet.getRow(2).getCell(i).toString();
            date2 = date2.replace("startDate", countryRequest.getStarDate().toString());
            date2 = date2.replace("endDate", countryRequest.getEndDate().toString());
            sheet.getRow(2).getCell(i).setCellValue(date2);
        }

        String nameFile = "";

        int rowLast = 0;

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {

            rowLast = sheet.getLastRowNum();
            XSSFRow row = sheet.createRow(rowLast + 1);
            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);
            XSSFCell number = row.createCell(0);
            number.setCellValue(cellCount + 1);
            number.setCellStyle(cellStyle);

            Import.createRowsImport(xssfWorkbook, row, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 5);

        }

        rowLast = sheet.getLastRowNum();
        XSSFRow rowTotal = sheet.createRow(rowLast + 1);

        Import.createRowsImport(xssfWorkbook, rowTotal, cellStyle, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн", 5);

        if (countryRequest.isImport()) {
            if (countryRequest.isProduct()) {
                fromTitle = fromTitle.replace("reqCountryOrProduct", countryRequest.getReqCountryOrProduct());
                fromTitle = fromTitle.replace("importexport", "поступлении в Республику Беларусь ");
                fromTitle2 = fromTitle2.replace("resCountryOrProduct", "Страна отправления");
                nameFile = "ImportProduct";
            } else {
                fromTitle = fromTitle.replace("reqCountryOrProduct", countryRequest.getReqCountryOrProduct());
                fromTitle = fromTitle.replace("importexport", "поступлении в Республику Беларусь из");
                fromTitle2 = fromTitle2.replace("resCountryOrProduct", "Наименование подкарантинной продукции");
                nameFile = "ImportCountry";
            }
            sheet.getRow(1).getCell(1).setCellValue(fromTitle2);
            sheet.getRow(0).getCell(0).setCellValue(fromTitle);
            CountryRow countryRow = countryRequest.getCountryRows().get(0);


        } else {

            if (countryRequest.isProduct()) {
                fromTitle = fromTitle.replace("reqCountryOrProduct", countryRequest.getReqCountryOrProduct());
                fromTitle = fromTitle.replace("importexport", "вывозе из Республики Беларусь ");
                fromTitle2 = fromTitle2.replace("resCountryOrProduct", "Страна получатель");
                nameFile = "ExportProduсt";
            } else {
                fromTitle = fromTitle.replace("reqCountryOrProduct", countryRequest.getReqCountryOrProduct());
                fromTitle = fromTitle.replace("importexport", "вывозе из Республики Беларусь в ");
                fromTitle2 = fromTitle2.replace("resCountryOrProduct", "Наименование подкарантинной продукции");
                nameFile = "ExportCountry";
            }
        }
        sheet.getRow(1).getCell(1).setCellValue(fromTitle2);
        sheet.getRow(0).getCell(0).setCellValue(fromTitle);

        CountryRow countryRow = countryRequest.getCountryRows().get(0);
        if (countryRequest.isFlowers()) {
            ReExport.createRowsMaterial(xssfWorkbook, countryRequest, cellStyle, 2);
        }
        Import.createRowsAllFss(xssfWorkbook, countryRequest, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 2);

        Import.createRowsNameObl(xssfWorkbook, cellStyle);

        Import.createRowsFss2022(xssfWorkbook, countryRequest, cellStyle, 2);

        try (
                OutputStream fileOut = new FileOutputStream("src/main/resources/" + nameFile + LocalDate.now() + ".xlsx")) {
            xssfWorkbook.write(fileOut);
        }

        return "";
    }

    @PostMapping("/re-export")
    public String printReData(@RequestBody CountryReport countryRequest) throws Exception {

        log.debug("CountryReport " + countryRequest);
        TotalMass totalMass = new TotalMass(countryRequest);
        log.debug("TotalMass " + totalMass);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook("src/main/resources/BlancRe.xlsx");
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStyle.setWrapText(true);

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        //font.setItalic(true);
        cellStyle.setFont(font);


        String fromTitle = sheet.getRow(0).getCell(0).toString();
        String fromTitle2 = sheet.getRow(1).getCell(1).toString();

        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

        for (int i = 2; i <= 17; i++) {
            String date2 = sheet.getRow(2).getCell(i).toString();
            date2 = date2.replace("startDate", countryRequest.getStarDate().toString());
            date2 = date2.replace("endDate", countryRequest.getEndDate().toString());
            sheet.getRow(2).getCell(i).setCellValue(date2);
        }
        String nameFile = "";

        int rowLast = 0;

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {

            rowLast = sheet.getLastRowNum();
            XSSFRow row = sheet.createRow(rowLast + 1);
            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);
            XSSFCell number = row.createCell(0);
            number.setCellValue(cellCount + 1);
            number.setCellStyle(cellStyle);

            Import.createRowsImport(xssfWorkbook, row, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 5);

        }

        rowLast = sheet.getLastRowNum();
        XSSFRow rowTotal = sheet.createRow(rowLast + 1);

        Import.createRowsImport(xssfWorkbook, rowTotal, cellStyle, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн", 5);

        if (countryRequest.isReexport()) {

            fromTitle = fromTitle.replace("countryExport", "в Российскую Федерацию");
            nameFile = "ReExportInRF";


        } else {
            fromTitle = fromTitle.replace("countryExport", "");
            nameFile = "ReExportAllCountry";

        }
        sheet.getRow(0).getCell(0).setCellValue(fromTitle);

        CountryRow countryRow = countryRequest.getCountryRows().get(0);
        if (countryRequest.isFlowers()) {
            ReExport.createRowsMaterial(xssfWorkbook, countryRequest, cellStyle, 2);
        }
        Import.createRowsAllFss(xssfWorkbook, countryRequest, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 2);

        Import.createRowsNameObl(xssfWorkbook, cellStyle);

        Import.createRowsFss2022(xssfWorkbook, countryRequest, cellStyle, 2);

        try (
                OutputStream fileOut = new FileOutputStream("src/main/resources/" + nameFile + LocalDate.now() + ".xlsx")) {
            xssfWorkbook.write(fileOut);
        }

        return "";
    }

    @PostMapping("/tranzit")
    public String printTranData(@RequestBody CountryReport countryRequest) throws Exception {

        log.debug("CountryReport " + countryRequest);
//        TotalMass totalMass = new TotalMass(countryRequest);
//        log.debug("TotalMass " + totalMass);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook("src/main/resources/BlancTranzit.xlsx");
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStyle.setWrapText(true);

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        //font.setItalic(true);
        cellStyle.setFont(font);


        String nameFile = "";

        int rowLast = 0;

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {

            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);
            ElementRegion elementRegion = countryRow.getRegions().get(cellCount);
            Tranzit.createRows(xssfWorkbook, cellStyle, countryRow.getRegions(), elementRegion.getNamePoints());

        }

        if (countryRequest.isTranzitEAEU()) {
            nameFile = "TranzitEAEUandCIS";
        } else {
            nameFile = "Tranzit";
        }

        try (
                OutputStream fileOut = new FileOutputStream("src/main/resources/" + nameFile + LocalDate.now() + ".xlsx")) {
            xssfWorkbook.write(fileOut);
        }

        return null;
    }


}

