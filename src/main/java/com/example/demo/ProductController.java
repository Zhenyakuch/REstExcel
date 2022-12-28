package com.example.demo;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.example.demo.model.TotalMass;
import com.example.demo.model.dto.CountryReport;
import com.example.demo.model.dto.CountryRow;
import com.example.demo.model.dto.ElementMass;
import com.example.demo.model.dto.ElementRegion;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

@RestController
@Slf4j
public class ProductController {

    @PostMapping("/product")
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
        String fromTitle2= sheet.getRow(1).getCell(1).toString();
        fromTitle = fromTitle.replace("startDate", countryRequest.getStarDate().toString());
        fromTitle = fromTitle.replace("endDate", countryRequest.getEndDate().toString());

        fromTitle = fromTitle.replace("reqCountryOrProduct", countryRequest.getReqCountryOrProduct());
        fromTitle2 = fromTitle2.replace("resCountryOrProduct", "Страна направления");

        sheet.getRow(1).getCell(1).setCellValue(fromTitle2);
        sheet.getRow(0).getCell(0).setCellValue(fromTitle);


        int rowLast = 0;
        for (int cellCount = 0; cellCount < countryRequest.getCountryRows().size(); cellCount++) {
            rowLast = sheet.getLastRowNum();
            XSSFRow row = sheet.createRow(rowLast + 1);

            CountryRow countryRow = countryRequest.getCountryRows().get(cellCount);

            XSSFCell numer = row.createCell(0);
            numer.setCellValue(cellCount +1);
            numer.setCellStyle(cellStyle);

            createRows(xssfWorkbook, row, countryRow.getRegions(), countryRow.getMassProduct(), countryRow.getResCountryOrProduct());
        }
        rowLast = sheet.getLastRowNum();
        XSSFRow rowTotal = sheet.createRow(rowLast + 1);
        //rowTotal = sheet.createRow(new CellRangeAddress(rowLast + 1,2));
//        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1,rowLast + 1,0,1));
       // System.out.println("rwwefwrre   "+ new CellRangeAddress(rowLast + 1,rowLast + 1,0,1).getNumberOfCells());

        createRows(xssfWorkbook, rowTotal, totalMass.getRegions(), totalMass.getMassProduct(), "ИТОГО, тонн");

        try (OutputStream fileOut = new FileOutputStream("src/main/resources/updated.xlsx")) {
            xssfWorkbook.write(fileOut);
        }

        return "";
    }

    private static void createRows(XSSFWorkbook xssfWorkbook, XSSFRow row, List<ElementRegion> regions, ElementMass mass, String country) {

        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setWrapText(true);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);


        XSSFCell cellCountry = row.createCell(1);
        cellCountry.setCellStyle(cellStyle);
        XSSFCell cellMassProductDateWeight = row.createCell(2);
        cellMassProductDateWeight.setCellStyle(cellStyle);
        XSSFCell cellMassProductWeekWeight = row.createCell(3);
        cellMassProductWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsBrestDateWeight = row.createCell(4);
        cellRegionsBrestDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsBrestWeekWeight = row.createCell(5);
        cellRegionsBrestWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsVitebskDateWeight = row.createCell(6);
        cellRegionsVitebskDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsVitebskWeekWeight = row.createCell(7);
        cellRegionsVitebskWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsGomelDateWeight = row.createCell(8);
        cellRegionsGomelDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsGomelWeekWeight = row.createCell(9);
        cellRegionsGomelWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsGrodnoDateWeight = row.createCell(10);
        cellRegionsGrodnoDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsGrodnoWeekWeight = row.createCell(11);
        cellRegionsGrodnoWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMinskDateWeight = row.createCell(12);
        cellRegionsMinskDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMinskWeekWeight = row.createCell(13);
        cellRegionsMinskWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMogilevDateWeight = row.createCell(14);
        cellRegionsMogilevDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMogilexWeekWeight = row.createCell(15);
        cellRegionsMogilexWeekWeight.setCellStyle(cellStyle);

        cellCountry.setCellValue(country);
        cellMassProductDateWeight.setCellValue(mass.dateWeightDouble());
        cellMassProductWeekWeight.setCellValue(mass.weekWeightDouble());

        for (int i = 0; regions.size() > i; i++) {
            ElementRegion elementRegion = regions.get(i);
            switch (elementRegion.getRegion()) {
                case 1:
                    cellRegionsBrestDateWeight.setCellValue(elementRegion.dateWeightDouble());
                    cellRegionsBrestWeekWeight.setCellValue(elementRegion.weekWeightDouble());
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