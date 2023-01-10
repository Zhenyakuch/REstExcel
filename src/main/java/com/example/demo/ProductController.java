package com.example.demo;

import com.example.demo.model.TotalMass;
import com.example.demo.model.dto.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.util.CellRangeAddress;
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

        String fromTitle = sheet.getRow(0).getCell(0).toString();
        String fromTitle2 = sheet.getRow(1).getCell(1).toString();

        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

        for (int i = 2; i <= 14; i++) {
            String date2 = sheet.getRow(2).getCell(i).toString();
            date2 = date2.replace("startDate", countryRequest.getStarDate().toString());
            date2 = date2.replace("endDate", countryRequest.getEndDate().toString());
            sheet.getRow(2).getCell(i).setCellValue(date2);
        }

        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

        String namefile = "";

        if (countryRequest.isImport()) {
            fromTitle = fromTitle.replace("reqCountryOrProduct", countryRequest.getReqCountryOrProduct());
            fromTitle = fromTitle.replace("importexport", "поступлении в Республику Беларусь ");
            fromTitle2 = fromTitle2.replace("resCountryOrProduct", "Страна направления");
            namefile = "import";
        } else {
            fromTitle = fromTitle.replace("reqCountryOrProduct", countryRequest.getReqCountryOrProduct());
            fromTitle = fromTitle.replace("importexport", "вывозе из Республики Беларусь ");
            fromTitle2 = fromTitle2.replace("resCountryOrProduct", "Наименование подкарантинной продукции");
            namefile = "export";
        }

        sheet.getRow(1).getCell(1).setCellValue(fromTitle2);
        sheet.getRow(0).getCell(0).setCellValue(fromTitle);

        int rowLast = 0;

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {
            rowLast = sheet.getLastRowNum();
            XSSFRow row = sheet.createRow(rowLast + 1);
            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);
            XSSFCell numer = row.createCell(0);
            numer.setCellValue(cellCount + 1);
            numer.setCellStyle(cellStyle);

            Import.createRowsImport(xssfWorkbook, row, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 3);
        }

        rowLast = sheet.getLastRowNum();
        XSSFRow rowTotal = sheet.createRow(rowLast + 1);

        Import.createRowsImport(xssfWorkbook, rowTotal, cellStyle, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн", 3);

        try (OutputStream fileOut = new FileOutputStream("src/main/resources/" + namefile + LocalDate.now() + ".xlsx")) {
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
        cellStyle.setWrapText(true);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        String fromTitle = sheet.getRow(0).getCell(0).toString();
        String fromTitle2 = sheet.getRow(1).getCell(1).toString();

        for (int i = 2; i <= 16; i++) {
            String date2 = sheet.getRow(2).getCell(i).toString();
            date2 = date2.replace("startDate", countryRequest.getStarDate().toString());
            date2 = date2.replace("endDate", countryRequest.getEndDate().toString());
            sheet.getRow(2).getCell(i).setCellValue(date2);
        }

        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

        String namefile = "";

        if (countryRequest.isImport()) {
            fromTitle = fromTitle.replace("importexport", "в РФ");
            namefile = "ReExportRF";
        } else {
            fromTitle = fromTitle.replace("importexport", "");
            namefile = "ReExport";
        }

        sheet.getRow(1).getCell(1).setCellValue(fromTitle2);
        sheet.getRow(0).getCell(0).setCellValue(fromTitle);

        int rowLast = 0;

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {
            rowLast = sheet.getLastRowNum();
            XSSFRow row = sheet.createRow(rowLast + 1);
            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);

            XSSFCell numer = row.createCell(0);
            numer.setCellValue(cellCount + 1);
            numer.setCellStyle(cellStyle);

            ReExport.createOneRows(xssfWorkbook, row, cellStyle, countryRow.getMassProduct(), 4);
            ReExport.createRows20212(xssfWorkbook, row, cellStyle, countryRow.getRegions(), 18);
            ReExport.createRowsReExport(xssfWorkbook, row, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 5);
        }

        rowLast = sheet.getLastRowNum();
        XSSFRow rowTotal = sheet.createRow(rowLast + 1);

        ReExport.createRowsReExport(xssfWorkbook, rowTotal, cellStyle, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн", 5);


        //
        CountryRow countryRow = countryRequest.getCountryRows().get(0);//проверить дебагом
        ReExport.createRowsMaterial(xssfWorkbook, countryRequest, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 2);

        ReExport.createRowsAllFss(xssfWorkbook, countryRequest, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 2);

        XSSFRow fssobl = sheet.createRow(rowLast + 5);
        XSSFCell fssobl0 = fssobl.createCell(0);
        XSSFCell fssobl1 = fssobl.createCell(1);
        XSSFCell fssobl2 = fssobl.createCell(2);
        XSSFCell fssobl3 = fssobl.createCell(3);
        XSSFCell fssobl4 = fssobl.createCell(4);
        XSSFCell fssobl5 = fssobl.createCell(5);
        fssobl0.setCellStyle(cellStyle);
        fssobl1.setCellStyle(cellStyle);
        fssobl2.setCellStyle(cellStyle);
        fssobl3.setCellStyle(cellStyle);
        fssobl4.setCellStyle(cellStyle);
        fssobl5.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 5, rowLast + 5, 0, 5));

        XSSFCell fssBrest = fssobl.createCell(6);
        XSSFCell fssBrest2 = fssobl.createCell(7);
        XSSFCell fssVitebsk = fssobl.createCell(8);
        XSSFCell fssVitebsk2 = fssobl.createCell(9);
        XSSFCell fssGomel = fssobl.createCell(10);
        XSSFCell fssGomel2 = fssobl.createCell(11);
        XSSFCell fssGrodno = fssobl.createCell(12);
        XSSFCell fssGrodno2 = fssobl.createCell(13);
        XSSFCell fssMinsk = fssobl.createCell(14);
        XSSFCell fssMinsk2 = fssobl.createCell(15);
        XSSFCell fssMogilev = fssobl.createCell(16);
        XSSFCell fssMogilev2 = fssobl.createCell(17);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 5, rowLast + 5, 6, 7));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 5, rowLast + 5, 8, 9));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 5, rowLast + 5, 10, 11));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 5, rowLast + 5, 12, 13));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 5, rowLast + 5, 14, 15));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 5, rowLast + 5, 16, 17));
        fssBrest.setCellValue("Брест");
        fssVitebsk.setCellValue("Витебск");
        fssGomel.setCellValue("Гомель");
        fssGrodno.setCellValue("Гродно");
        fssMinsk.setCellValue("Минск");
        fssMogilev.setCellValue("Могилев");
        fssBrest.setCellStyle(cellStyle);
        fssVitebsk.setCellStyle(cellStyle);
        fssGomel.setCellStyle(cellStyle);
        fssGrodno.setCellStyle(cellStyle);
        fssMinsk.setCellStyle(cellStyle);
        fssMogilev.setCellStyle(cellStyle);
        fssBrest2.setCellStyle(cellStyle);
        fssVitebsk2.setCellStyle(cellStyle);
        fssGomel2.setCellStyle(cellStyle);
        fssGrodno2.setCellStyle(cellStyle);
        fssMinsk2.setCellStyle(cellStyle);
        fssMogilev2.setCellStyle(cellStyle);


        ReExport.createRowsFss2021(xssfWorkbook, countryRequest, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 2);
        ReExport.createRowsFss2022(xssfWorkbook, countryRequest, cellStyle, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 2);

        try (OutputStream fileOut = new FileOutputStream("src/main/resources/" + namefile + LocalDate.now() + ".xlsx")) {
            xssfWorkbook.write(fileOut);
        }

        return "";
    }

    @PostMapping("/tranzit")
    public String printTranData(@RequestBody CountryReport countryRequest) throws Exception {

        log.debug("CountryReport " + countryRequest);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook("src/main/resources/BlancTranzit.xlsx");
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);

        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        String namefile = "Tranzit";

        int rowLast = 0;

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {
            rowLast = sheet.getLastRowNum();
            XSSFRow row = sheet.createRow(rowLast + 1);
            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);

            Tranzit.createRows(xssfWorkbook, row, cellStyle, countryRow.getRegions());
        }

//        rowLast = sheet.getLastRowNum();
//        XSSFRow rowTotal = sheet.createRow(rowLast + 1);
//
//        Import.createRowsImport(xssfWorkbook, rowTotal, cellStyle, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн", 3);

        try (OutputStream fileOut = new FileOutputStream("src/main/resources/" + namefile + LocalDate.now() + ".xlsx")) {
            xssfWorkbook.write(fileOut);
        }
        return null;
    }


}

