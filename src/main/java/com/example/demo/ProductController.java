package com.example.demo;

import com.example.demo.model.TotalMass;
import com.example.demo.model.dto.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;

@RestController
@Slf4j
public class ProductController {
    String nameFile = "";
    int rowLast = 0;

    @PostMapping("/import-export")
    public String printData(@RequestBody CountryReport countryRequest) throws Exception {

        log.debug("CountryReport " + countryRequest);
        TotalMass totalMass = new TotalMass(countryRequest);
        log.debug("TotalMass " + totalMass);

        OPCPackage pkg = OPCPackage.open("src/main/resources/Blanc.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(pkg);

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
        cellStyle.setFont(font);

        XSSFCellStyle cellStyleRow = xssfWorkbook.createCellStyle();
        cellStyleRow.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setWrapText(true);
        cellStyleRow.setAlignment(CellStyle.ALIGN_CENTER);

        XSSFFont fontRow = xssfWorkbook.createFont();
        fontRow.setFontHeightInPoints((short) 12);
        fontRow.setFontName("Times New Roman");
        cellStyleRow.setFont(font);

        String fromTitle = sheet.getRow(0).getCell(0).toString();
        String fromTitle2 = sheet.getRow(1).getCell(1).toString();

        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

        for (int i = 2; i <= 17; i++) {

            String date2 = sheet.getRow(2).getCell(i).toString();
            date2 = date2.replace("startDate", countryRequest.getStarDate().toString());
            date2 = date2.replace("endDate", countryRequest.getEndDate().toString());

            if (countryRequest.isImport()) {
                date2 = date2.replace("importexport2", "поступило с");

            } else {
                date2 = date2.replace("importexport2", "вывезено с");

            }
            sheet.getRow(2).getCell(i).setCellValue(date2);
        }

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {

            rowLast = sheet.getLastRowNum();
            XSSFRow row = sheet.createRow(rowLast + 1);
            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);
            XSSFCell number = row.createCell(0);
            number.setCellValue(cellCount + 1);
            number.setCellStyle(cellStyleRow);

            Import.createRowsImport(xssfWorkbook, row, cellStyle, cellStyleRow, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 5);

        }

        rowLast = sheet.getLastRowNum();
        XSSFRow rowTotal = sheet.createRow(rowLast + 1);

        Import.createRowsImport(xssfWorkbook, rowTotal, cellStyle, cellStyleRow, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн", 5);

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

        if (countryRequest.isFlowers()) {
            ReExport.createRowsMaterial(xssfWorkbook, countryRequest, cellStyle, cellStyleRow, 2);
        }
        if(countryRequest.getFss()!=null) {
            Import.createRowsAllFss(xssfWorkbook, countryRequest, cellStyle, cellStyleRow, 2);

            Import.createRowsNameObl(xssfWorkbook, cellStyleRow);

            Import.createRowsFss2022(xssfWorkbook, countryRequest, cellStyle, cellStyleRow, 2);
        }

        String nameFileResponse = "src/main/resources/" + nameFile + LocalDate.now() + ".xlsx";

        try (
                OutputStream fileOut = Files.newOutputStream(Paths.get(nameFileResponse))) {
            xssfWorkbook.write(fileOut);
        }

        return ConBase64.convert(nameFileResponse);
    }

    @PostMapping("/re-export")
    public String printReData(@RequestBody CountryReport countryRequest) throws Exception {

        log.debug("CountryReport " + countryRequest);
        TotalMass totalMass = new TotalMass(countryRequest);
        log.debug("TotalMass " + totalMass);

        OPCPackage pkg = OPCPackage.open("src/main/resources/BlancRe.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(pkg);

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
        cellStyle.setFont(font);

        XSSFCellStyle cellStyleRow = xssfWorkbook.createCellStyle();
        cellStyleRow.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setWrapText(true);
        cellStyleRow.setAlignment(CellStyle.ALIGN_CENTER);

        XSSFFont fontRow = xssfWorkbook.createFont();
        fontRow.setFontHeightInPoints((short) 12);
        fontRow.setFontName("Times New Roman");
        cellStyleRow.setFont(font);

        String fromTitle = sheet.getRow(0).getCell(0).toString();

        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

        for (int i = 2; i <= 17; i++) {
            String date2 = sheet.getRow(2).getCell(i).toString();
            date2 = date2.replace("startDate", countryRequest.getStarDate().toString());
            date2 = date2.replace("endDate", countryRequest.getEndDate().toString());
            sheet.getRow(2).getCell(i).setCellValue(date2);
        }

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {

            rowLast = sheet.getLastRowNum();
            XSSFRow row = sheet.createRow(rowLast + 1);
            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);
            XSSFCell number = row.createCell(0);
            number.setCellValue(cellCount + 1);
            number.setCellStyle(cellStyleRow);

            Import.createRowsImport(xssfWorkbook, row, cellStyle, cellStyleRow, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(), 5);

        }

        rowLast = sheet.getLastRowNum();
        XSSFRow rowTotal = sheet.createRow(rowLast + 1);

        Import.createRowsImport(xssfWorkbook, rowTotal, cellStyle, cellStyleRow, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн", 5);

        if (countryRequest.isReexport()) {

            fromTitle = fromTitle.replace("countryExport", "в Российскую Федерацию");
            nameFile = "ReExportInRF";


        } else {
            fromTitle = fromTitle.replace("countryExport", "");
            nameFile = "ReExportAllCountry";

        }
        sheet.getRow(0).getCell(0).setCellValue(fromTitle);

        if (countryRequest.isFlowers()) {
            ReExport.createRowsMaterial(xssfWorkbook, countryRequest, cellStyle, cellStyleRow, 2);
        }
        Import.createRowsAllFss(xssfWorkbook, countryRequest, cellStyle, cellStyleRow, 2);

        Import.createRowsNameObl(xssfWorkbook, cellStyle);

        Import.createRowsFss2022(xssfWorkbook, countryRequest, cellStyle, cellStyleRow, 2);

        String nameFileResponse = "src/main/resources/" + nameFile + LocalDate.now() + ".xlsx";


        try (
                OutputStream fileOut = Files.newOutputStream(Paths.get(nameFileResponse))) {
            xssfWorkbook.write(fileOut);
        }

        return ConBase64.convert(nameFileResponse);
    }

    @PostMapping("/tranzit")
    public String printTranData(@RequestBody CountryReport countryRequest) throws Exception {

        log.debug("CountryReport " + countryRequest);

        OPCPackage pkg = OPCPackage.open("src/main/resources/BlancTranzit.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(pkg);

        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStyle.setWrapText(true);

        XSSFFont font = xssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setFontName("Times New Roman");
        cellStyle.setFont(font);

        XSSFCellStyle cellStyleRow = xssfWorkbook.createCellStyle();
        cellStyleRow.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStyleRow.setWrapText(true);
        cellStyleRow.setAlignment(CellStyle.ALIGN_CENTER);

        XSSFFont fontRow = xssfWorkbook.createFont();
        fontRow.setFontHeightInPoints((short) 12);
        fontRow.setFontName("Times New Roman");
        cellStyleRow.setFont(font);

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);

        String fromTitle = sheet.getRow(0).getCell(0).toString();

        for (int i = 1; i <= 11; i++) {

            String date2 = sheet.getRow(2).getCell(i).toString();
            String date1 = sheet.getRow(1).getCell(0).toString();

            if (countryRequest.isTranzitEAEU()) {
                fromTitle = fromTitle.replace("contents", "ТРАНЗИТ В АДРЕС СТРАН ЕВРАЗИЙСКОГО ЭКОНОМИЧЕСКОГО СОЮЗА И ГОСУДАРСТВ - УЧАСТНИЦ СНГ");
                date1 = date1.replace("name", "Наименование страны-получателя подкарантинной продукции");
                date2 = date2.replace("amount1", "тыс. т");
                date2 = date2.replace("amount2", "тыс. пос. ед.");
                date2 = date2.replace("amount3", "тыс. шт.");
                date2 = date2.replace("amount4", "тыс. парт.");
                date2 = date2.replace("amount5", "тыс. пак.");
                date2 = date2.replace("amount6", "м3");
            } else {
                fromTitle = fromTitle.replace("contents", "ТРАНЗИТ ПОДКАРАНТИННОЙ ПРОДУКЦИИ");
                date1 = date1.replace("name", "Наименование пограничных пунктов");
                date2 = date2.replace("amount1", "тыс. т");
                date2 = date2.replace("amount2", "тыс. пос. ед.");
                date2 = date2.replace("amount3", "тыс. шт.");
                date2 = date2.replace("amount4", "тыс. парт.");
                date2 = date2.replace("amount5", "тыс. м2");
                date2 = date2.replace("amount6", "тыс. м3");

            }

            sheet.getRow(2).getCell(i).setCellValue(date2);
            sheet.getRow(1).getCell(0).setCellValue(date1);
            sheet.getRow(0).getCell(0).setCellValue(fromTitle);
        }

        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().get(0).getRegions().size(); cellCount++) {

            if (countryRequest.isTranzitEAEU()) {

                nameFile = "TranzitEAEUandCIS";
                rowLast = sheet.getLastRowNum();
                XSSFRow row = sheet.createRow(rowLast + 1);
                CountryRow countryRow = countryRequest.getCountryRows().get(0);
                ElementRegion elementRegion = countryRow.getRegions().get(cellCount);
                Tranzit.create_obl(cellCount, xssfWorkbook, row, cellStyle, cellStyleRow, countryRow.getRegions(), elementRegion.getNamePoints());
                //В том числе в страны ЕАЭС
                Tranzit.plus_eaeu(countryRequest,countryRow,xssfWorkbook, sheet, cellStyleRow);

            } else {

                nameFile = "Tranzit";
                rowLast = sheet.getLastRowNum();
                XSSFRow row = sheet.createRow(rowLast + 1);
                CountryRow countryRow = countryRequest.getCountryRows().get(0);
                ElementRegion elementRegion = countryRow.getRegions().get(cellCount);
                Tranzit.create_obl(cellCount, xssfWorkbook, row, cellStyle, cellStyleRow, countryRow.getRegions(), elementRegion.getNamePoints());

            }

        }

        if (countryRequest.isTranzitEAEU() == false) {
            Tranzit.plusAll(xssfWorkbook, sheet, cellStyleRow,
                    Tranzit.summa_tonn2, Tranzit.summa_pos_ed2, Tranzit.summa_sht2, Tranzit.summa_part2, Tranzit.summa_m22,
                    Tranzit.summa_m32, Tranzit.summa_wagons2, Tranzit.summa_transport2, Tranzit.summa_container2,
                    Tranzit.summa_baggage2, Tranzit.summa_airplane2);
        }


        String nameFileResponse = "src/main/resources/" + nameFile + LocalDate.now() + ".xlsx";

        try (
                OutputStream fileOut = Files.newOutputStream(Paths.get(nameFileResponse))) {
            xssfWorkbook.write(fileOut);
        }

        return ConBase64.convert(nameFileResponse);
    }


}

