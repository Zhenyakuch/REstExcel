package com.example.demo.model.dto;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class Import {

    public static void createRowsImport(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow, List<ElementRegion> regions, ElementMass mass, String country, int i) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast33 = sheet.getLastRowNum();
        XSSFCell cellCountry = row.createCell(1);
        XSSFCell cellCountry2 = row.createCell(2);
        cellCountry.setCellStyle(cellStyle);
        cellCountry2.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast33, rowLast33, 1, 2));
        cellCountry.setCellValue(country);

        XSSFCell cellMassProductDateFromTo = row.createCell(3); // дата с по
        cellMassProductDateFromTo.setCellStyle(cellStyleRow);
        XSSFCell cellMassProductWeekWeight = row.createCell(4);// данные за неделю
        cellMassProductWeekWeight.setCellStyle(cellStyleRow);
        XSSFCell cellMassProductDateWeight = row.createCell(i);// дата с 01.01.2022
        cellMassProductDateWeight.setCellStyle(cellStyleRow);

        XSSFCell cellRegionsBrestDateWeight = row.createCell(i + 1);
        cellRegionsBrestDateWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsBrestWeekWeight = row.createCell(i + 2);
        cellRegionsBrestWeekWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsVitebskDateWeight = row.createCell(i + 3);
        cellRegionsVitebskDateWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsVitebskWeekWeight = row.createCell(i + 4);
        cellRegionsVitebskWeekWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsGomelDateWeight = row.createCell(i + 5);
        cellRegionsGomelDateWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsGomelWeekWeight = row.createCell(i + 6);
        cellRegionsGomelWeekWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsGrodnoDateWeight = row.createCell(i + 7);
        cellRegionsGrodnoDateWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsGrodnoWeekWeight = row.createCell(i + 8);
        cellRegionsGrodnoWeekWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsMinskDateWeight = row.createCell(i + 9);
        cellRegionsMinskDateWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsMinskWeekWeight = row.createCell(i + 10);
        cellRegionsMinskWeekWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsMogilevDateWeight = row.createCell(i + 11);
        cellRegionsMogilevDateWeight.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsMogilexWeekWeight = row.createCell(i + 12);
        cellRegionsMogilexWeekWeight.setCellStyle(cellStyleRow);

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

        fssPeriod2023.setCellStyle(cellStyleRow);
        fssPeriod7.setCellStyle(cellStyleRow);
        fssPeriod2022.setCellStyle(cellStyleRow);
        cellRegionsBrest2022.setCellStyle(cellStyleRow);
        cellRegionsBrest2022_7.setCellStyle(cellStyleRow);
        cellRegionsVitebsk2022.setCellStyle(cellStyleRow);
        cellRegionsVitebsk2022_7.setCellStyle(cellStyleRow);
        cellRegionsGomel2022.setCellStyle(cellStyleRow);
        cellRegionsGomel2022_7.setCellStyle(cellStyleRow);
        cellRegionsGrodno2022.setCellStyle(cellStyleRow);
        cellRegionsGrodno2022_7.setCellStyle(cellStyleRow);
        cellRegionsMinsk2022.setCellStyle(cellStyleRow);
        cellRegionsMinsk2022_7.setCellStyle(cellStyleRow);
        cellRegionsMogilev2022.setCellStyle(cellStyleRow);
        cellRegionsMogilev2022_7.setCellStyle(cellStyleRow);

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

        XSSFCell fssPeriod2023 = row.createCell(i + 1);
        fssPeriod2023.setCellStyle(cellStyleRow);
        XSSFCell fssPeriod7 = row.createCell(i + 2);
        fssPeriod7.setCellStyle(cellStyleRow);
        XSSFCell fssPeriod2022 = row.createCell(i + 3);
        fssPeriod2022.setCellStyle(cellStyleRow);

        XSSFCell cellRegionsBrest = row.createCell(i + 4);
        cellRegionsBrest.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsBrest_7 = row.createCell(i + 5);
        cellRegionsBrest_7.setCellStyle(cellStyleRow);

        XSSFCell cellRegionsVitebsk = row.createCell(i + 6);
        cellRegionsVitebsk.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsVitebsk_7 = row.createCell(i + 7);
        cellRegionsVitebsk_7.setCellStyle(cellStyleRow);

        XSSFCell cellRegionsGomel = row.createCell(i + 8);
        cellRegionsGomel.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsGomel_7 = row.createCell(i + 9);
        cellRegionsGomel_7.setCellStyle(cellStyleRow);

        XSSFCell cellRegionsGrodno = row.createCell(i + 10);
        cellRegionsGrodno.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsGrodno_7 = row.createCell(i + 11);
        cellRegionsGrodno_7.setCellStyle(cellStyleRow);

        XSSFCell cellRegionsMinsk = row.createCell(i + 12);
        cellRegionsMinsk.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsMinsk_7 = row.createCell(i + 13);
        cellRegionsMinsk_7.setCellStyle(cellStyleRow);

        XSSFCell cellRegionsMogilev = row.createCell(i + 14);
        cellRegionsMogilev.setCellStyle(cellStyleRow);
        XSSFCell cellRegionsMogilev_7 = row.createCell(i + 15);
        cellRegionsMogilev_7.setCellStyle(cellStyleRow);


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

        XSSFRow fss_obl = sheet.createRow(rowLast + 1);
        XSSFCell fss_obl0 = fss_obl.createCell(0);
        XSSFCell fss_obl1 = fss_obl.createCell(1);
        XSSFCell fss_obl2 = fss_obl.createCell(2);
        XSSFCell fss_obl3 = fss_obl.createCell(3);
        XSSFCell fss_obl4 = fss_obl.createCell(4);
        XSSFCell fss_obl5 = fss_obl.createCell(5);
        fss_obl0.setCellStyle(cellStyle);
        fss_obl1.setCellStyle(cellStyle);
        fss_obl2.setCellStyle(cellStyle);
        fss_obl3.setCellStyle(cellStyle);
        fss_obl4.setCellStyle(cellStyle);
        fss_obl5.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 5));

        XSSFCell fssBrest = fss_obl.createCell(6);
        XSSFCell fssBrest2 = fss_obl.createCell(7);
        XSSFCell fssVitebsk = fss_obl.createCell(8);
        XSSFCell fssVitebsk2 = fss_obl.createCell(9);
        XSSFCell fssGomel = fss_obl.createCell(10);
        XSSFCell fssGomel2 = fss_obl.createCell(11);
        XSSFCell fssGrodno = fss_obl.createCell(12);
        XSSFCell fssGrodno2 = fss_obl.createCell(13);
        XSSFCell fssMinsk = fss_obl.createCell(14);
        XSSFCell fssMinsk2 = fss_obl.createCell(15);
        XSSFCell fssMogilev = fss_obl.createCell(16);
        XSSFCell fssMogilev2 = fss_obl.createCell(17);
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
