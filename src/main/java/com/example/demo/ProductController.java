package com.example.demo;

import com.aspose.cells.Cells;
import com.aspose.cells.ReplaceOptions;
import com.example.demo.model.TotalMass;
import com.example.demo.model.dto.CountryReport;
import com.example.demo.model.dto.CountryRow;
import com.example.demo.model.dto.ElementMass;
import com.example.demo.model.dto.ElementRegion;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.time.LocalDate;
import java.util.List;

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
        cellStyle.setWrapText(true);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        String fromTitle = sheet.getRow(0).getCell(0).toString();
         String fromTitle2 = sheet.getRow(1).getCell(1).toString();

        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

         for (int i=2; i<=14; i++){
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
            fromTitle = fromTitle.replace("importexport", "поступлении в Республику Беларусь");
            fromTitle2 = fromTitle2.replace("resCountryOrProduct", "Страна направления");
            namefile = "import";
        }else {
            fromTitle = fromTitle.replace("reqCountryOrProduct", countryRequest.getReqCountryOrProduct());
            fromTitle = fromTitle.replace("importexport", "вывозе из Республики Беларусь");
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

            createRowsImport(xssfWorkbook, row, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(),3);
        }
        rowLast = sheet.getLastRowNum();
        XSSFRow rowTotal = sheet.createRow(rowLast + 1);

        createRowsImport(xssfWorkbook, rowTotal, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн",3);

        try (OutputStream fileOut = new FileOutputStream("src/main/resources/" + namefile + LocalDate.now() + ".xlsx")) {
            xssfWorkbook.write(fileOut);
        }

        return "";
    }


    @PostMapping("/re-export")
    public String printReData(@RequestBody CountryReport countryRequest) throws Exception{

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

//        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
//        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

        for (int i=2; i<=16; i++){
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
        }else {
            fromTitle = fromTitle.replace("importexport", "");
            namefile = "ReExport";
        }

        sheet.getRow(1).getCell(1).setCellValue(fromTitle2);
        sheet.getRow(0).getCell(0).setCellValue(fromTitle);


        int rowLast = 0;
        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {
            rowLast = sheet.getLastRowNum();
            XSSFRow row = sheet.createRow(rowLast +1);

            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);

            XSSFCell numer = row.createCell(0);
            numer.setCellValue(cellCount + 1);
            numer.setCellStyle(cellStyle);

            createOneRows(xssfWorkbook,row,countryRow.getMassProduct(),4);
            createRows20212(xssfWorkbook,row,countryRow.getRegions(),18);
            createRowsReExport(xssfWorkbook, row, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct(),5);
        }
        rowLast = sheet.getLastRowNum();
        XSSFRow rowTotal = sheet.createRow(rowLast + 1);

        createRowsReExport(xssfWorkbook, rowTotal, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн",5);
//
        XSSFRow rowTotal2 = sheet.createRow(rowLast + 2);
        XSSFCell cellCountry = rowTotal2.createCell(0);
        XSSFCell cellCountry2 = rowTotal2.createCell(1);


        cellCountry.setCellStyle(cellStyle);
        cellCountry2.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast+2,rowLast+2,0,1));
//
        XSSFRow rowTotal3 = sheet.createRow(rowLast + 3);
        XSSFCell cellCountry3 = rowTotal3.createCell(0);
        XSSFCell cellCountry33 = rowTotal3.createCell(1);

        cellCountry3.setCellStyle(cellStyle);
        cellCountry33.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast+3,rowLast+3,0,1));

       // sheet.addMergedRegion(new CellRangeAddress(rowLast+2,rowLast+3,0,0));

       cellCountry3.setCellValue("Срезы и посадочный материал цветочной и лесодекоративной, горшечной продукции");
       rowTotal3.setHeight((short) 1100);
        XSSFCell massa = rowTotal2.createCell(2);
        massa.setCellStyle(cellStyle);
        massa.setCellValue("тр. ед.");
        XSSFCell massa2 = rowTotal3.createCell(2);
        massa2.setCellStyle(cellStyle);
        massa2.setCellValue("млн. шт.");

        XSSFRow rowTotal4 = sheet.createRow(rowLast + 4);
        XSSFCell fss = rowTotal4.createCell(0);
        XSSFCell fss2 = rowTotal4.createCell(1);
        XSSFCell fss3 = rowTotal4.createCell(2);
        fss.setCellStyle(cellStyle);
        fss2.setCellStyle(cellStyle);
        fss3.setCellStyle(cellStyle);
        rowTotal4.setHeight((short) 900);
        sheet.addMergedRegion(new CellRangeAddress(rowLast+4,rowLast+4,0,2));
        fss.setCellValue("Выдано ФСС, оформленных на реэкспорт (на всю продукцию), шт");

        XSSFRow rowTotal6 = sheet.createRow(rowLast + 6);
        XSSFCell fss2021 = rowTotal6.createCell(0);
        XSSFCell fss22021 = rowTotal6.createCell(1);
        XSSFCell fss32021 = rowTotal6.createCell(2);
        fss2021.setCellStyle(cellStyle);
        fss22021.setCellStyle(cellStyle);
        fss32021.setCellStyle(cellStyle);
        rowTotal6.setHeight((short) 900);
        sheet.addMergedRegion(new CellRangeAddress(rowLast+6,rowLast+6,0,2));
        fss2021.setCellValue("Выдано ФСС, оформленных на реэкспорт             (на всю продукцию), шт в 2021 г.");


        XSSFRow rowTotal7 = sheet.createRow(rowLast + 7);
        XSSFCell fss2022 = rowTotal7.createCell(0);
        XSSFCell fss22022 = rowTotal7.createCell(1);
        XSSFCell fss32022 = rowTotal7.createCell(2);
        fss2022.setCellStyle(cellStyle);
        fss22022.setCellStyle(cellStyle);
        fss32022.setCellStyle(cellStyle);
        rowTotal7.setHeight((short) 900);
        sheet.addMergedRegion(new CellRangeAddress(rowLast+7,rowLast+7,0,2));
        fss2022.setCellValue("Выдано ФСС, оформленных на реэкспорт             (на всю продукцию), шт в 2022 г.");

        //




        try (OutputStream fileOut = new FileOutputStream("src/main/resources/" + namefile + LocalDate.now() + ".xlsx")) {
            xssfWorkbook.write(fileOut);
        }

        return "";
    }

    private static void createRowsReExport(XSSFWorkbook xssfWorkbook, XSSFRow row, List<ElementRegion> regions, ElementMass mass, String country, int i) {

        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setWrapText(true);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);


        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast33 = sheet.getLastRowNum();
        //XSSFRow rowTotal = sheet.createRow(rowLast33 + 1);
        XSSFCell cellCountry = row.createCell(1);
        XSSFCell cellCountry2 = row.createCell(2);
        cellCountry.setCellStyle(cellStyle);
        cellCountry2.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast33,rowLast33,1,2));
        cellCountry.setCellValue(country);
        //rowTotal2.setHeight((short) 1100);



        XSSFCell cellMassProductDateWeight = row.createCell(3);
        cellMassProductDateWeight.setCellStyle(cellStyle);
        XSSFCell  cellMassProductWeekWeight = row.createCell(i);
        cellMassProductWeekWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsBrestDateWeight = row.createCell(i+1);
        cellRegionsBrestDateWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsBrestWeekWeight = row.createCell(i+2);
        cellRegionsBrestWeekWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsVitebskDateWeight = row.createCell(i+3);
        cellRegionsVitebskDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsVitebskWeekWeight = row.createCell(i+4);
        cellRegionsVitebskWeekWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsGomelDateWeight = row.createCell(i+5);
        cellRegionsGomelDateWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsGomelWeekWeight = row.createCell(i+6);
        cellRegionsGomelWeekWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsGrodnoDateWeight = row.createCell(i+7);
        cellRegionsGrodnoDateWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsGrodnoWeekWeight = row.createCell(i+8);
        cellRegionsGrodnoWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMinskDateWeight = row.createCell(i+9);
        cellRegionsMinskDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMinskWeekWeight = row.createCell(i+10);
        cellRegionsMinskWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMogilevDateWeight = row.createCell(i+11);
        cellRegionsMogilevDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMogilexWeekWeight = row.createCell(i+12);
        cellRegionsMogilexWeekWeight.setCellStyle(cellStyle);
//        }


        cellMassProductDateWeight.setCellValue(mass.dateWeightDouble());
        cellMassProductWeekWeight.setCellValue(mass.weekWeightDouble());


//        XSSFCell  cellRegionsBrestDateWeight2021 = row.createCell(17);
//        CountryReport countryRequest = new CountryReport();

        for (int k = 0; regions.size() > k; k++) {
            ElementRegion elementRegion = regions.get(k);
            switch (elementRegion.getRegion()) {
                case 1:
                    cellRegionsBrestDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsBrestWeekWeight.setCellValue(elementRegion.weekWeightDouble());
//                    cellRegionsBrestDateWeight2021.setCellValue("21dfv");
                    break;
                case 2:
                    cellRegionsVitebskDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsVitebskWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;
                case 3:
                    cellRegionsGomelDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsGomelWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;
                case 4:
                    cellRegionsGrodnoDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsGrodnoWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;
                case 5:
                    cellRegionsMinskDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsMinskWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;
                case 6:
                    cellRegionsMogilevDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsMogilexWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;

            }
        }
    }

    private static void createOneRows(XSSFWorkbook xssfWorkbook, XSSFRow row, ElementMass mass, int i){
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setWrapText(true);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);


        XSSFCell cellMassProductDateFromTo = row.createCell(i);
        cellMassProductDateFromTo.setCellStyle(cellStyle);
        cellMassProductDateFromTo.setCellValue(mass.dateFromToDouble());

    }

    private static void createRows20212(XSSFWorkbook xssfWorkbook, XSSFRow row, List<ElementRegion> region, int i){
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setWrapText(true);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFCell  cellBrest2021= row.createCell(i);
        cellBrest2021.setCellStyle(cellStyle);
        cellBrest2021.setCellValue(region.get(0).getDate2021());
        XSSFCell  cellVitebsk2021 = row.createCell(i+1);
        cellVitebsk2021.setCellStyle(cellStyle);
        cellVitebsk2021.setCellValue(region.get(1).getDate2021());
        XSSFCell  cellGomel2021 = row.createCell(i+2);
        cellGomel2021.setCellStyle(cellStyle);
        cellGomel2021.setCellValue(region.get(2).getDate2021());
        XSSFCell cellGrodno2021 = row.createCell(i+3);
        cellGrodno2021.setCellStyle(cellStyle);
        cellGrodno2021.setCellValue(region.get(3).getDate2021());
        XSSFCell  cellMinsk2021 = row.createCell(i+4);
        cellMinsk2021.setCellStyle(cellStyle);
        cellMinsk2021.setCellValue(region.get(4).getDate2021());
        XSSFCell  cellMogilev2021 = row.createCell(i+5);
        cellMogilev2021.setCellStyle(cellStyle);
        cellMogilev2021.setCellValue(region.get(5).getDate2021());

        XSSFCell  cellBrest2022= row.createCell(i+6);
        cellBrest2022.setCellStyle(cellStyle);
        cellBrest2022.setCellValue(region.get(0).getDate2022());
        XSSFCell  cellVitebsk2022 = row.createCell(i+7);
        cellVitebsk2022.setCellStyle(cellStyle);
        cellVitebsk2022.setCellValue(region.get(1).getDate2022());
        XSSFCell  cellGomel2022 = row.createCell(i+8);
        cellGomel2022.setCellStyle(cellStyle);
        cellGomel2022.setCellValue(region.get(2).getDate2022());
        XSSFCell cellGrodno2022 = row.createCell(i+9);
        cellGrodno2022.setCellStyle(cellStyle);
        cellGrodno2022.setCellValue(region.get(3).getDate2022());
        XSSFCell  cellMinsk2022 = row.createCell(i+10);
        cellMinsk2022.setCellStyle(cellStyle);
        cellMinsk2022.setCellValue(region.get(4).getDate2022());
        XSSFCell  cellMogilev2022 = row.createCell(i+11);
        cellMogilev2022.setCellStyle(cellStyle);
        cellMogilev2022.setCellValue(region.get(5).getDate2022());

//        cellMassProductDateFromTo.setCellValue(mass.dateFromToDouble());

    }

    private static void createRowsImport(XSSFWorkbook xssfWorkbook, XSSFRow row, List<ElementRegion> regions, ElementMass mass, String country, int i) {

        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setWrapText(true);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);

//
        XSSFCell cellCountry = row.createCell(1);
            cellCountry.setCellStyle(cellStyle);
        XSSFCell cellMassProductDateWeight = row.createCell(2);
            cellMassProductDateWeight.setCellStyle(cellStyle);
        XSSFCell  cellMassProductWeekWeight = row.createCell(i);
            cellMassProductWeekWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsBrestDateWeight = row.createCell(i+1);
            cellRegionsBrestDateWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsBrestWeekWeight = row.createCell(i+2);
            cellRegionsBrestWeekWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsVitebskDateWeight = row.createCell(i+3);
            cellRegionsVitebskDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsVitebskWeekWeight = row.createCell(i+4);
            cellRegionsVitebskWeekWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsGomelDateWeight = row.createCell(i+5);
            cellRegionsGomelDateWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsGomelWeekWeight = row.createCell(i+6);
            cellRegionsGomelWeekWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsGrodnoDateWeight = row.createCell(i+7);
            cellRegionsGrodnoDateWeight.setCellStyle(cellStyle);
        XSSFCell  cellRegionsGrodnoWeekWeight = row.createCell(i+8);
            cellRegionsGrodnoWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMinskDateWeight = row.createCell(i+9);
            cellRegionsMinskDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMinskWeekWeight = row.createCell(i+10);
            cellRegionsMinskWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMogilevDateWeight = row.createCell(i+11);
            cellRegionsMogilevDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMogilexWeekWeight = row.createCell(i+12);
            cellRegionsMogilexWeekWeight.setCellStyle(cellStyle);
//        }

        cellCountry.setCellValue(country);
        cellMassProductDateWeight.setCellValue(mass.dateWeightDouble());
        cellMassProductWeekWeight.setCellValue(mass.weekWeightDouble());


//        XSSFCell  cellRegionsBrestDateWeight2021 = row.createCell(17);
//        CountryReport countryRequest = new CountryReport();

        for (int k = 0; regions.size() > k; k++) {
            ElementRegion elementRegion = regions.get(k);
            switch (elementRegion.getRegion()) {
                case 1:
                    cellRegionsBrestDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsBrestWeekWeight.setCellValue(elementRegion.weekWeightDouble());
//                    cellRegionsBrestDateWeight2021.setCellValue("21dfv");
                    break;
                case 2:
                    cellRegionsVitebskDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsVitebskWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;
                case 3:
                    cellRegionsGomelDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsGomelWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;
                case 4:
                    cellRegionsGrodnoDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsGrodnoWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;
                case 5:
                    cellRegionsMinskDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsMinskWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;
                case 6:
                    cellRegionsMogilevDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsMogilexWeekWeight.setCellValue(elementRegion.weekWeightDouble());
                    break;

            }
        }
    }


}

