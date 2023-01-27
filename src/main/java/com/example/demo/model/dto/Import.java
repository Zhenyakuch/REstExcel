package com.example.demo.model.dto;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class Import {

    public static void createRowsImport(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow, List<ElementRegion> regions, ElementMass mass, String country, int i) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast33 = sheet.getLastRowNum();
        XSSFCell cellCountry = getXssfCell(row, 1, cellStyle);
        XSSFCell cellCountry2 = getXssfCell(row, 2, cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast33, rowLast33, 1, 2));
        cellCountry.setCellValue(country);

        XSSFCell cellMassProductDateFromTo = getXssfCell(row, 3, cellStyleRow);
        XSSFCell cellMassProductWeekWeight = getXssfCell(row, 4, cellStyleRow);
        XSSFCell cellMassProductDateWeight = getXssfCell(row, i, cellStyleRow);

        XSSFCell cellRegionsBrestDateWeight = getXssfCell(row, i + 1, cellStyleRow);
        XSSFCell cellRegionsBrestWeekWeight = getXssfCell(row, i + 2, cellStyleRow);
        XSSFCell cellRegionsVitebskDateWeight = getXssfCell(row, i + 3, cellStyleRow);
        XSSFCell cellRegionsVitebskWeekWeight = getXssfCell(row, i + 4, cellStyleRow);
        XSSFCell cellRegionsGomelDateWeight = getXssfCell(row, i + 5, cellStyleRow);
        XSSFCell cellRegionsGomelWeekWeight = getXssfCell(row, i + 6, cellStyleRow);
        XSSFCell cellRegionsGrodnoDateWeight = getXssfCell(row, i + 7, cellStyleRow);
        XSSFCell cellRegionsGrodnoWeekWeight = getXssfCell(row, i + 8, cellStyleRow);
        XSSFCell cellRegionsMinskDateWeight = getXssfCell(row, i + 9, cellStyleRow);
        XSSFCell cellRegionsMinskWeekWeight = getXssfCell(row, i + 10, cellStyleRow);
        XSSFCell cellRegionsMogilevDateWeight = getXssfCell(row, i + 11, cellStyleRow);
        XSSFCell cellRegionsMogilexWeekWeight = getXssfCell(row, i + 12, cellStyleRow);

        cellCountry.setCellValue(country);
        cellMassProductDateWeight.setCellValue(mass.dateWeightDouble());
        cellMassProductWeekWeight.setCellValue(mass.weekWeightDouble());
        cellMassProductDateFromTo.setCellValue(mass.dateFromToDouble());


        for (ElementRegion elementRegion : regions) {
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


    public static void createRowsFss2022(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow, int i) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast = sheet.getLastRowNum();

        XSSFRow row = sheet.createRow(rowLast + 1);
        XSSFCell fss2022 = row.createCell(0);
        XSSFCell fss22022 = row.createCell(1);
        XSSFCell fss32022 = row.createCell(2);
        fss2022.setCellStyle(cellStyle);
        fss22022.setCellStyle(cellStyle);
        fss32022.setCellStyle(cellStyle);
        row.setHeight((short) 900);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 2));
        fss2022.setCellValue("Выдано ФСС, (на всю продукцию), шт в 2022 г.");

        XSSFCell fssPeriod2023 = getXssfCell(row, i + 1, cellStyleRow);
        XSSFCell fssPeriod7 = getXssfCell(row, i + 2, cellStyleRow);
        XSSFCell fssPeriod2022 = getXssfCell(row, i + 3, cellStyleRow);
        XSSFCell cellRegionsBrest = getXssfCell(row, i + 4, cellStyleRow);
        XSSFCell cellRegionsBrest_7 = getXssfCell(row, i + 5, cellStyleRow);
        XSSFCell cellRegionsVitebsk = getXssfCell(row, i + 6, cellStyleRow);
        XSSFCell cellRegionsVitebsk_7 = getXssfCell(row, i + 7, cellStyleRow);
        XSSFCell cellRegionsGomel = getXssfCell(row, i + 8, cellStyleRow);
        XSSFCell cellRegionsGomel_7 = getXssfCell(row, i + 9, cellStyleRow);
        XSSFCell cellRegionsGrodno = getXssfCell(row, i + 10, cellStyleRow);
        XSSFCell cellRegionsGrodno_7 = getXssfCell(row, i + 11, cellStyleRow);
        XSSFCell cellRegionsMinsk = getXssfCell(row, i + 12, cellStyleRow);
        XSSFCell cellRegionsMinsk_7 = getXssfCell(row, i + 13, cellStyleRow);
        XSSFCell cellRegionsMogilev = getXssfCell(row, i + 14, cellStyleRow);
        XSSFCell cellRegionsMogilev_7 = getXssfCell(row, i + 15, cellStyleRow);

        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, i + 1, i + 15));

        fssPeriod2023.setCellValue(countryRequest.getAllFss2022());

    }


    public static void createRowsAllFss(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow,  int i) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast = sheet.getLastRowNum();

        XSSFRow row = sheet.createRow(rowLast + 1);
        XSSFCell all_fss = row.createCell(0);
        XSSFCell all_fss2 = row.createCell(1);
        XSSFCell all_fss3 = row.createCell(2);
        all_fss.setCellStyle(cellStyle);
        all_fss2.setCellStyle(cellStyle);
        all_fss3.setCellStyle(cellStyle);
        row.setHeight((short) 900);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 2));
        all_fss.setCellValue("Выдано ФСС, (на всю продукцию), шт");

        XSSFCell fssPeriod2023 = getXssfCell(row, i + 1, cellStyleRow);
        XSSFCell fssPeriod7 = getXssfCell(row, i + 2, cellStyleRow);
        XSSFCell fssPeriod2022 = getXssfCell(row, i + 3, cellStyleRow);
        XSSFCell cellRegionsBrest = getXssfCell(row, i + 4, cellStyleRow);
        XSSFCell cellRegionsBrest_7 = getXssfCell(row, i + 5, cellStyleRow);
        XSSFCell cellRegionsVitebsk = getXssfCell(row, i + 6, cellStyleRow);
        XSSFCell cellRegionsVitebsk_7 = getXssfCell(row, i + 7, cellStyleRow);
        XSSFCell cellRegionsGomel = getXssfCell(row, i + 8, cellStyleRow);
        XSSFCell cellRegionsGomel_7 = getXssfCell(row, i + 9, cellStyleRow);
        XSSFCell cellRegionsGrodno = getXssfCell(row, i + 10, cellStyleRow);
        XSSFCell cellRegionsGrodno_7 = getXssfCell(row, i + 11, cellStyleRow);
        XSSFCell cellRegionsMinsk = getXssfCell(row, i + 12, cellStyleRow);
        XSSFCell cellRegionsMinsk_7 = getXssfCell(row, i + 13, cellStyleRow);
        XSSFCell cellRegionsMogilev = getXssfCell(row, i + 14, cellStyleRow);
        XSSFCell cellRegionsMogilev_7 = getXssfCell(row, i + 15, cellStyleRow);

        fssPeriod2023.setCellValue(countryRequest.getFss().getAllFss());
        fssPeriod7.setCellValue(countryRequest.getFss().getAllFss_7());
        fssPeriod2022.setCellValue("***");

        cellRegionsBrest.setCellValue(countryRequest.getFss().getAllFssBrest());
        cellRegionsVitebsk.setCellValue(countryRequest.getFss().getAllFssVitebsk());
        cellRegionsGomel.setCellValue(countryRequest.getFss().getAllFssGomel());
        cellRegionsGrodno.setCellValue(countryRequest.getFss().getAllFssGrodno());
        cellRegionsMinsk.setCellValue(countryRequest.getFss().getAllFssMinsk());
        cellRegionsMogilev.setCellValue(countryRequest.getFss().getAllFssMogilev());

        cellRegionsBrest_7.setCellValue(countryRequest.getFss().getAllFss_7Brest());
        cellRegionsVitebsk_7.setCellValue(countryRequest.getFss().getAllFss_7Vitebsk());
        cellRegionsGomel_7.setCellValue(countryRequest.getFss().getAllFss_7Gomel());
        cellRegionsGrodno_7.setCellValue(countryRequest.getFss().getAllFss_7Grodno());
        cellRegionsMinsk_7.setCellValue(countryRequest.getFss().getAllFss_7Minsk());
        cellRegionsMogilev_7.setCellValue(countryRequest.getFss().getAllFss_7Mogilev());


    }

    public static void createRowsNameObl(XSSFWorkbook xssfWorkbook, XSSFCellStyle cellStyle) {
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast = sheet.getLastRowNum();

        XSSFRow row = sheet.createRow(rowLast + 1);
        XSSFCell fss_obl0 = getXssfCell(row, 0, cellStyle);
        XSSFCell fss_obl1 = getXssfCell(row, 1, cellStyle);
        XSSFCell fss_obl2 = getXssfCell(row, 2, cellStyle);
        XSSFCell fss_obl3 = getXssfCell(row, 3, cellStyle);
        XSSFCell fss_obl4 = getXssfCell(row, 4, cellStyle);
        XSSFCell fss_obl5 = getXssfCell(row, 5, cellStyle);

        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 5));

        XSSFCell fssBrest = getXssfCell(row, 6, cellStyle);
        XSSFCell fssBrest2 = getXssfCell(row, 7, cellStyle);
        XSSFCell fssVitebsk = getXssfCell(row, 8, cellStyle);
        XSSFCell fssVitebsk2 =getXssfCell(row, 9, cellStyle);
        XSSFCell fssGomel = getXssfCell(row, 10, cellStyle);
        XSSFCell fssGomel2 = getXssfCell(row, 11, cellStyle);
        XSSFCell fssGrodno = getXssfCell(row, 12, cellStyle);
        XSSFCell fssGrodno2 = getXssfCell(row, 13, cellStyle);
        XSSFCell fssMinsk = getXssfCell(row, 14, cellStyle);
        XSSFCell fssMinsk2 = getXssfCell(row, 15, cellStyle);
        XSSFCell fssMogilev = getXssfCell(row, 16, cellStyle);
        XSSFCell fssMogilev2 =getXssfCell(row, 17, cellStyle);

        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 6, 7));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 8, 9));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 10, 11));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 12, 13));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 14, 15));
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 16, 17));

        fssBrest.setCellValue("Брест");
        fssVitebsk.setCellValue("Витебск");
        fssGomel.setCellValue("Гомель");
        fssGrodno.setCellValue("Гродно");
        fssMinsk.setCellValue("Минск");
        fssMogilev.setCellValue("Могилев");
    }

    private static XSSFCell getXssfCell(XSSFRow row, int i, XSSFCellStyle cellStyleRow) {
        XSSFCell cellRegionsBrestDateWeight = row.createCell(i);
        cellRegionsBrestDateWeight.setCellStyle(cellStyleRow);
        return cellRegionsBrestDateWeight;
    }
}
