package com.example.demo;

import com.example.demo.model.TotalMass;
import com.example.demo.model.dto.*;
import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.*;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.nio.file.Files;


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

        InputStream is = getClass().getClassLoader().getResourceAsStream("Blanc.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);

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
            Import.createRowsImport(xssfWorkbook, row, cellStyle, cellStyleRow, countryRow.getRegions(), countryRow.getMassProduct(),
                    countryRow.getResCountryOrProduct(), 5);
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

        if (countryRequest.getFss() != null) {
            Import.createRowsAllFss(xssfWorkbook, countryRequest, cellStyle, cellStyleRow, 2);
            Import.createRowsNameObl(xssfWorkbook, cellStyleRow);
            Import.createRowsFss2022(xssfWorkbook, countryRequest, cellStyle, cellStyleRow, 2);
        }

        File tempFile = File.createTempFile(nameFile, null);
        try (OutputStream fileOut = Files.newOutputStream(tempFile.toPath())) {
            xssfWorkbook.write(fileOut);
        }

        return ConBase64.convert(tempFile);
    }

    @PostMapping("/re-export")
    public String printReData(@RequestBody CountryReport countryRequest) throws Exception {

        log.debug("CountryReport " + countryRequest);
        TotalMass totalMass = new TotalMass(countryRequest);
        log.debug("TotalMass " + totalMass);

        InputStream is = getClass().getClassLoader().getResourceAsStream("BlancRe.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);

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

        File tempFile = File.createTempFile(nameFile, null);

        try (OutputStream fileOut = Files.newOutputStream(tempFile.toPath())) {
            xssfWorkbook.write(fileOut);
        }

        return ConBase64.convert(tempFile);
    }

    @PostMapping("/tranzit")
    public String printTranData(@RequestBody CountryReport countryRequest) throws Exception {

        log.debug("CountryReport " + countryRequest);

        InputStream is = getClass().getClassLoader().getResourceAsStream("BlancTranzit.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);

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

        XSSFCellStyle cell_styl_obl = xssfWorkbook.createCellStyle();
        cell_styl_obl.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cell_styl_obl.setWrapText(true);//перенос слов

        XSSFFont font2 = xssfWorkbook.createFont();
        font2.setFontHeightInPoints((short) 12);
        font2.setFontName("Times New Roman");
        font2.setBold(true);
        cell_styl_obl.setFont(font2);

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
                Tranzit.create_obl(cellCount, xssfWorkbook, row, cellStyle, cell_styl_obl, cellStyleRow, countryRow.getRegions(), elementRegion.getNamePoints());
                //В том числе в страны ЕАЭС
                Tranzit.plus_eaeu(countryRequest, countryRow, xssfWorkbook, sheet, cellStyleRow, cell_styl_obl);
            } else {
                nameFile = "Tranzit";
                rowLast = sheet.getLastRowNum();
                XSSFRow row = sheet.createRow(rowLast + 1);
                CountryRow countryRow = countryRequest.getCountryRows().get(0);
                ElementRegion elementRegion = countryRow.getRegions().get(cellCount);
                Tranzit.create_obl(cellCount, xssfWorkbook, row, cellStyle, cell_styl_obl, cellStyleRow, countryRow.getRegions(), elementRegion.getNamePoints());
            }
        }

        if (!countryRequest.isTranzitEAEU()) {
            Tranzit.plusAll(xssfWorkbook, sheet, cellStyleRow,
                    Tranzit.summa_tonn2, Tranzit.summa_pos_ed2, Tranzit.summa_sht2, Tranzit.summa_part2, Tranzit.summa_m22,
                    Tranzit.summa_m32, Tranzit.summa_wagons2, Tranzit.summa_transport2, Tranzit.summa_container2,
                    Tranzit.summa_baggage2, Tranzit.summa_airplane2);
        }

        Tranzit.nullable();
        File tempFile = File.createTempFile(nameFile, null);

        try (OutputStream fileOut = Files.newOutputStream(tempFile.toPath())) {
            xssfWorkbook.write(fileOut);
        }

        return ConBase64.convert(tempFile);
    }

    @PostMapping("/sticker")
    public String createPdfSticker(@RequestBody Sticker sticker) {

        String nameFile = "ЭТИКЕТКА";
        try {

            // InputStream is = getClass().getClassLoader().getResourceAsStream("Sticker.docx");

            Document document = new Document("C:\\Users\\Evgeniya.Kychinskaya\\Desktop\\Belfito Project\\src\\main\\resources\\Sticker.docx");

            // Replace a specific text
            document.replace("number", String.valueOf(sticker.getNumber()), true, true);
            document.replace("name", sticker.getName(), true, true);
            document.replace("weight", String.valueOf(sticker.getWeight()), true, true);
            document.replace("origin", sticker.getOrigin(), true, true);
            document.replace("place", sticker.getPlace(), true, true);
            document.replace("net_weight", String.valueOf(sticker.getNet_weight()), true, true);
            document.replace("recipient", sticker.getRecipient(), true, true);
            document.replace("appointment", sticker.getAppointment(), true, true);
            document.replace("area", String.valueOf(sticker.getArea()), true, true);
            document.replace("external_sings", sticker.getExternal_sings(), true, true);
            document.replace("provisional_definition", sticker.getProvisional_definition(), true, true);
            document.replace("additional_info", sticker.getAdditional_info(), true, true);
            document.replace("seal_number", sticker.getSeal_number(), true, true);
            document.replace("position", sticker.getPosition(), true, true);
            document.replace("date", String.valueOf(sticker.getDate()), true, true);
            document.replace("FIO1", sticker.getFio1(), true, true);
            document.replace("FIO2", sticker.getFio2(), true, true);

            //Save the result document
            document.saveToFile(nameFile, FileFormat.Docx);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return convert(nameFile);
    }

    @PostMapping("/conclusion")
    public String createPdfConclusion(@RequestBody Conclusion conclusion) {

        String nameFile = "ЗАКЛЮЧЕНИЕ";
        try {
            Document document = new Document("C:\\Users\\Evgeniya.Kychinskaya\\Desktop\\Belfito Project\\src\\main\\resources\\Conclusion.docx");
            // Replace a specific text
            document.replace("name_legal", conclusion.getName_legal(), true, true);
            document.replace("date1", String.valueOf(conclusion.getDate1()), true, true);
            document.replace("date2", String.valueOf(conclusion.getDate2()), true, true);
            document.replace("date3", String.valueOf(conclusion.getDate3()), true, true);
            document.replace("date4", String.valueOf(conclusion.getDate4()), true, true);
            document.replace("number1", String.valueOf(conclusion.getNumber1()), true, true);
            document.replace("number2", String.valueOf(conclusion.getNumber2()), true, true);
            document.replace("number3", String.valueOf(conclusion.getNumber3()), true, true);
            document.replace("issued", conclusion.getIssued(), true, true);
            document.replace("name", conclusion.getName(), true, true);
            document.replace("weight", String.valueOf(conclusion.getWeight()), true, true);
            document.replace("origin", conclusion.getOrigin(), true, true);
            document.replace("place", conclusion.getPlace(), true, true);
            document.replace("from_whos", conclusion.getFrom_whos(), true, true);
            document.replace("recipient", conclusion.getRecipient(), true, true);
            document.replace("result", conclusion.getResult(), true, true);
            document.replace("events", conclusion.getEvents(), true, true);
            document.replace("FIO", conclusion.getFio(), true, true);

            //Save the result document
            document.saveToFile(nameFile, FileFormat.Docx);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return convert(nameFile);
    }

    @PostMapping("/act-disinfection")
    public String createPdfActDecontamination(@RequestBody Disinfection disinfection) {

        String nameFile = "Акт обеззараживания";
        try {
            Document document = new Document("C:\\Users\\Evgeniya.Kychinskaya\\Desktop\\Belfito Project\\src\\main\\resources\\ActDisinfection.docx");
            // Replace a specific text
            document.replace("date1", String.valueOf(disinfection.getDate1()), true, true);
            document.replace("date2", String.valueOf(disinfection.getDate2()), true, true);
            document.replace("number", String.valueOf(disinfection.getNumber()), true, true);
            document.replace("name1", disinfection.getName1(), true, true);
            document.replace("name2", disinfection.getName2(), true, true);
            document.replace("quantity", String.valueOf(disinfection.getQuantity()), true, true);
            document.replace("conclusion1", disinfection.getConclusion1(), true, true);
            document.replace("conclusion2", disinfection.getConclusion2(), true, true);
            document.replace("conclusion3", disinfection.getConclusion3(), true, true);
            document.replace("organization", disinfection.getOrganization(), true, true);
            document.replace("method_disinfection", disinfection.getMethod_disinfection(), true, true);
            document.replace("FIO1", disinfection.getFio1(), true, true);
            document.replace("FIO2", disinfection.getFio2(), true, true);
            document.replace("FIO3", disinfection.getFio3(), true, true);

            //Save the result document
            document.saveToFile(nameFile, FileFormat.Docx);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return convert(nameFile);
    }

    @PostMapping("/act-destruction")
    public String createPdfActDestruction(@RequestBody Destruction destruction) {

        String nameFile = "Акт об уничтожении";
        try {
            Document document = new Document("C:\\Users\\Evgeniya.Kychinskaya\\Desktop\\Belfito Project\\src\\main\\resources\\ActDestruction.docx");

            document.replace("number", destruction.getName(), true, true);
            document.replace("date1", String.valueOf(destruction.getDate1()), true, true);
            document.replace("date2", String.valueOf(destruction.getDate2()), true, true);
            document.replace("method_destruction", destruction.getMethod_destruction(), true, true);
            document.replace("name", destruction.getName(), true, true);
            document.replace("quantity", String.valueOf(destruction.getQuantity()), true, true);
            document.replace("weight", String.valueOf(destruction.getWeight()), true, true);
            document.replace("place", destruction.getPlace(), true, true);
            document.replace("position1", destruction.getPosition1(), true, true);
            document.replace("position2", destruction.getPosition2(), true, true);
            document.replace("position3", destruction.getPosition3(), true, true);
            document.replace("FIO1", destruction.getFio1(), true, true);
            document.replace("FIO2", destruction.getFio2(), true, true);
            document.replace("FIO3", destruction.getFio3(), true, true);

            //Save the result document
            document.saveToFile(nameFile, FileFormat.Docx);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return convert(nameFile);
    }

    @PostMapping("/act-return")
    public String createPdfActReturn() {

        String nameFile = "Акт возврата";
        try {
            Document document = new Document("C:\\Users\\Evgeniya.Kychinskaya\\Desktop\\Belfito Project\\src\\main\\resources\\ActReturn.docx");

            document.replace("number", "33333", true, true);
            document.replace("data1", "04.02.2023", true, true);
            document.replace("place", "гаражи", true, true);
            document.replace("FIO1", "Гек О.К.", true, true);
            document.replace("FIO2", "Чук О.К.", true, true);
            document.replace("FIO3", "Перец О.К.", true, true);
            document.replace("name", "яблочки", true, true);
            document.replace("quantity", "1000000", true, true);
            document.replace("recipient", "ОАО арарара", true, true);
            document.replace("place_sender", "порт", true, true);
            document.replace("number_TS", "4567АГ-7", true, true);
            document.replace("numberFSS", "34635636", true, true);
            document.replace("data2", "05.02.2023", true, true);
            document.replace("return_reasons", "зараженный товар", true, true);
            document.replace("organizationFSS", "БЕЛФИТО", true, true);

            //Save the result document
            document.saveToFile(nameFile, FileFormat.Docx);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return convert(nameFile);
    }

    public String convert(String nameFile) {
        try {
            InputStream templateInputStream = new FileInputStream(nameFile);
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(templateInputStream);
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

            String outputfilepath = "C:\\Users\\Evgeniya.Kychinskaya\\Desktop\\Belfito Project\\src\\main\\resources\\" + nameFile + ".pdf";
            FileOutputStream filePdf = new FileOutputStream(outputfilepath);
            Docx4J.toPDF(wordMLPackage, filePdf);
            filePdf.flush();
            filePdf.close();

        } catch (Throwable e) {
            e.printStackTrace();
        }

        return nameFile;
    }


}