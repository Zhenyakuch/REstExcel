package com.example.demo.model.dto;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class Import {

    public static void createRowsImport(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, List<ElementRegion> regions, ElementMass mass, String country, int i) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast33 = sheet.getLastRowNum();
        XSSFCell cellCountry = row.createCell(1);
        XSSFCell cellCountry2 = row.createCell(2);
        cellCountry.setCellStyle(cellStyle);
        cellCountry2.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast33, rowLast33, 1, 2));
        cellCountry.setCellValue(country);

        XSSFCell cellMassProsuctDateFromTo = row.createCell(3); // дата с по
        cellMassProsuctDateFromTo.setCellStyle(cellStyle);
        XSSFCell cellMassProductWeekWeight = row.createCell(4);// данные за неделю
        cellMassProductWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellMassProductDateWeight = row.createCell(i);// дата с 01.01.2022
        cellMassProductDateWeight.setCellStyle(cellStyle);

        XSSFCell cellRegionsBrestDateWeight = row.createCell(i + 1);
        cellRegionsBrestDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsBrestWeekWeight = row.createCell(i + 2);
        cellRegionsBrestWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsVitebskDateWeight = row.createCell(i + 3);
        cellRegionsVitebskDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsVitebskWeekWeight = row.createCell(i + 4);
        cellRegionsVitebskWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsGomelDateWeight = row.createCell(i + 5);
        cellRegionsGomelDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsGomelWeekWeight = row.createCell(i + 6);
        cellRegionsGomelWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsGrodnoDateWeight = row.createCell(i + 7);
        cellRegionsGrodnoDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsGrodnoWeekWeight = row.createCell(i + 8);
        cellRegionsGrodnoWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMinskDateWeight = row.createCell(i + 9);
        cellRegionsMinskDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMinskWeekWeight = row.createCell(i + 10);
        cellRegionsMinskWeekWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMogilevDateWeight = row.createCell(i + 11);
        cellRegionsMogilevDateWeight.setCellStyle(cellStyle);
        XSSFCell cellRegionsMogilexWeekWeight = row.createCell(i + 12);
        cellRegionsMogilexWeekWeight.setCellStyle(cellStyle);

        cellCountry.setCellValue(country);
        cellMassProductDateWeight.setCellValue(mass.dateWeightDouble());
        cellMassProductWeekWeight.setCellValue(mass.weekWeightDouble());
        cellMassProsuctDateFromTo.setCellValue(mass.dateFromToDouble());


        for (int k = 0; regions.size() > k; k++) {
            ElementRegion elementRegion = regions.get(k);
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

    public static void createRowsFss2022(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, int i) {

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

        XSSFCell fssPeriod2023 = row.createCell(i + 1);
        XSSFCell fssPeriod7 = row.createCell(i + 2);
        XSSFCell fssPeriod2022 = row.createCell(i + 3);
        XSSFCell cellRegionsBrest2022 = row.createCell(i + 4);
        XSSFCell cellRegionsBrest2022_7 = row.createCell(i + 5);
        XSSFCell cellRegionsVitebsk2022 = row.createCell(i + 6);
        XSSFCell cellRegionsVitebsk2022_7 = row.createCell(i + 7);
        XSSFCell cellRegionsGomel2022 = row.createCell(i + 8);
        XSSFCell cellRegionsGomel2022_7 = row.createCell(i + 9);
        XSSFCell cellRegionsGrodno2022 = row.createCell(i + 10);
        XSSFCell cellRegionsGrodno2022_7 = row.createCell(i + 11);
        XSSFCell cellRegionsMinsk2022 = row.createCell(i + 12);
        XSSFCell cellRegionsMinsk2022_7 = row.createCell(i + 13);
        XSSFCell cellRegionsMogilev2022 = row.createCell(i + 14);
        XSSFCell cellRegionsMogilev2022_7 = row.createCell(i + 15);

        fssPeriod2023.setCellStyle(cellStyle);
        fssPeriod7.setCellStyle(cellStyle);
        fssPeriod2022.setCellStyle(cellStyle);
        cellRegionsBrest2022.setCellStyle(cellStyle);
        cellRegionsBrest2022_7.setCellStyle(cellStyle);
        cellRegionsVitebsk2022.setCellStyle(cellStyle);
        cellRegionsVitebsk2022_7.setCellStyle(cellStyle);
        cellRegionsGomel2022.setCellStyle(cellStyle);
        cellRegionsGomel2022_7.setCellStyle(cellStyle);
        cellRegionsGrodno2022.setCellStyle(cellStyle);
        cellRegionsGrodno2022_7.setCellStyle(cellStyle);
        cellRegionsMinsk2022.setCellStyle(cellStyle);
        cellRegionsMinsk2022_7.setCellStyle(cellStyle);
        cellRegionsMogilev2022.setCellStyle(cellStyle);
        cellRegionsMogilev2022_7.setCellStyle(cellStyle);

        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, i + 1, i + 15));

        fssPeriod2023.setCellValue(countryRequest.getAllFss2022());


    }

    public static void createRowsAllFss(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, List<ElementRegion> regions, ElementMass mass, String country, int i) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast = sheet.getLastRowNum();

        XSSFRow row = sheet.createRow(rowLast + 1);
        XSSFCell allfss = row.createCell(0);
        XSSFCell allfss2 = row.createCell(1);
        XSSFCell allfss3 = row.createCell(2);
        allfss.setCellStyle(cellStyle);
        allfss2.setCellStyle(cellStyle);
        allfss3.setCellStyle(cellStyle);
        row.setHeight((short) 900);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 2));
        allfss.setCellValue("Выдано ФСС, (на всю продукцию), шт");

        XSSFCell fssPeriod2023 = row.createCell(i + 1);
        fssPeriod2023.setCellStyle(cellStyle);
        XSSFCell fssPeriod7 = row.createCell(i + 2);
        fssPeriod7.setCellStyle(cellStyle);
        XSSFCell fssPeriod2022 = row.createCell(i + 3);
        fssPeriod2022.setCellStyle(cellStyle);

        XSSFCell cellRegionsBrest = row.createCell(i + 4);
        cellRegionsBrest.setCellStyle(cellStyle);
        XSSFCell cellRegionsBrest_7 = row.createCell(i + 5);
        cellRegionsBrest_7.setCellStyle(cellStyle);

        XSSFCell cellRegionsVitebsk = row.createCell(i + 6);
        cellRegionsVitebsk.setCellStyle(cellStyle);
        XSSFCell cellRegionsVitebsk_7 = row.createCell(i + 7);
        cellRegionsVitebsk_7.setCellStyle(cellStyle);

        XSSFCell cellRegionsGomel = row.createCell(i + 8);
        cellRegionsGomel.setCellStyle(cellStyle);
        XSSFCell cellRegionsGomel_7 = row.createCell(i + 9);
        cellRegionsGomel_7.setCellStyle(cellStyle);

        XSSFCell cellRegionsGrodno = row.createCell(i + 10);
        cellRegionsGrodno.setCellStyle(cellStyle);
        XSSFCell cellRegionsGrodno_7 = row.createCell(i + 11);
        cellRegionsGrodno_7.setCellStyle(cellStyle);

        XSSFCell cellRegionsMinsk = row.createCell(i + 12);
        cellRegionsMinsk.setCellStyle(cellStyle);
        XSSFCell cellRegionsMinsk_7 = row.createCell(i + 13);
        cellRegionsMinsk_7.setCellStyle(cellStyle);

        XSSFCell cellRegionsMogilev = row.createCell(i + 14);
        cellRegionsMogilev.setCellStyle(cellStyle);
        XSSFCell cellRegionsMogilev_7 = row.createCell(i + 15);
        cellRegionsMogilev_7.setCellStyle(cellStyle);


        fssPeriod2023.setCellValue(countryRequest.getAllFss());
        fssPeriod7.setCellValue(countryRequest.getAllFss_7());
        fssPeriod2022.setCellValue("***");

        cellRegionsBrest.setCellValue(countryRequest.getAllFssBrest());
        cellRegionsVitebsk.setCellValue(countryRequest.getAllFssVitebsk());
        cellRegionsGomel.setCellValue(countryRequest.getAllFssGomel());
        cellRegionsGrodno.setCellValue(countryRequest.getAllFssGrodno());
        cellRegionsMinsk.setCellValue(countryRequest.getAllFssMinsk());
        cellRegionsMogilev.setCellValue(countryRequest.getAllFssMogilev());

        cellRegionsBrest_7.setCellValue(countryRequest.getAllFss_7Brest());
        cellRegionsVitebsk_7.setCellValue(countryRequest.getAllFss_7Vitebsk());
        cellRegionsGomel_7.setCellValue(countryRequest.getAllFss_7Gomel());
        cellRegionsGrodno_7.setCellValue(countryRequest.getAllFss_7Grodno());
        cellRegionsMinsk_7.setCellValue(countryRequest.getAllFss_7Minsk());
        cellRegionsMogilev_7.setCellValue(countryRequest.getAllFss_7Mogilev());


    }

    public static void createRowsNameObl(XSSFWorkbook xssfWorkbook, XSSFCellStyle cellStyle) {
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast = sheet.getLastRowNum();

        XSSFRow fssobl = sheet.createRow(rowLast + 1);
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
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 5));

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
    }
}
