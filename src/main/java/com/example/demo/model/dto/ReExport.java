package com.example.demo.model.dto;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class ReExport {

//    public static void createRowsReExport(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, List<ElementRegion> regions, ElementMass mass, String country, int i) {
//
//        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
//        int rowLast33 = sheet.getLastRowNum();
//        XSSFCell cellCountry = row.createCell(1);
//        XSSFCell cellCountry2 = row.createCell(2);
//        cellCountry.setCellStyle(cellStyle);
//        cellCountry2.setCellStyle(cellStyle);
//        sheet.addMergedRegion(new CellRangeAddress(rowLast33, rowLast33, 1, 2));
//        cellCountry.setCellValue(country);
//
//        XSSFCell cellMassProductDateWeight = row.createCell(3);
//        cellMassProductDateWeight.setCellStyle(cellStyle);
//        XSSFCell cellMassProductDateFromTo = row.createCell(4);
//        cellMassProductDateFromTo.setCellStyle(cellStyle);
//        XSSFCell cellMassProductWeekWeight = row.createCell(i);
//        cellMassProductWeekWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsBrestDateWeight = row.createCell(i + 1);
//        cellRegionsBrestDateWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsBrestWeekWeight = row.createCell(i + 2);
//        cellRegionsBrestWeekWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsVitebskDateWeight = row.createCell(i + 3);
//        cellRegionsVitebskDateWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsVitebskWeekWeight = row.createCell(i + 4);
//        cellRegionsVitebskWeekWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGomelDateWeight = row.createCell(i + 5);
//        cellRegionsGomelDateWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGomelWeekWeight = row.createCell(i + 6);
//        cellRegionsGomelWeekWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGrodnoDateWeight = row.createCell(i + 7);
//        cellRegionsGrodnoDateWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGrodnoWeekWeight = row.createCell(i + 8);
//        cellRegionsGrodnoWeekWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMinskDateWeight = row.createCell(i + 9);
//        cellRegionsMinskDateWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMinskWeekWeight = row.createCell(i + 10);
//        cellRegionsMinskWeekWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMogilevDateWeight = row.createCell(i + 11);
//        cellRegionsMogilevDateWeight.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMogilexWeekWeight = row.createCell(i + 12);
//        cellRegionsMogilexWeekWeight.setCellStyle(cellStyle);
//
//
//        XSSFCell cellBrest2021 = row.createCell(i + 13);
//        cellBrest2021.setCellStyle(cellStyle);
//        XSSFCell cellVitebsk2021 = row.createCell(i + 14);
//        cellVitebsk2021.setCellStyle(cellStyle);
//        XSSFCell cellGomel2021 = row.createCell(i + 15);
//        cellGomel2021.setCellStyle(cellStyle);
//        XSSFCell cellGrodno2021 = row.createCell(i + 16);
//        cellGrodno2021.setCellStyle(cellStyle);
//        XSSFCell cellMinsk2021 = row.createCell(i + 17);
//        cellMinsk2021.setCellStyle(cellStyle);
//        XSSFCell cellMogilev2021 = row.createCell(i + 18);
//        cellMogilev2021.setCellStyle(cellStyle);
//
//        XSSFCell cellBrest2022 = row.createCell(i + 19);
//        cellBrest2022.setCellStyle(cellStyle);
//        XSSFCell cellVitebsk2022 = row.createCell(i + 20);
//        cellVitebsk2022.setCellStyle(cellStyle);
//        XSSFCell cellGomel2022 = row.createCell(i + 21);
//        cellGomel2022.setCellStyle(cellStyle);
//        XSSFCell cellGrodno2022 = row.createCell(i + 22);
//        cellGrodno2022.setCellStyle(cellStyle);
//        XSSFCell cellMinsk2022 = row.createCell(i + 23);
//        cellMinsk2022.setCellStyle(cellStyle);
//        XSSFCell cellMogilev2022 = row.createCell(i + 24);
//        cellMogilev2022.setCellStyle(cellStyle);
//
//        cellMassProductDateWeight.setCellValue(mass.dateWeightDouble());
//        cellMassProductWeekWeight.setCellValue(mass.weekWeightDouble());
//        cellMassProductDateFromTo.setCellValue(mass.dateFromToDouble());
//
//        for (int k = 0; regions.size() > k; k++) {
//            ElementRegion elementRegion = regions.get(k);
//            switch (elementRegion.getRegion()) {
//                case 1:
//                    cellRegionsBrestDateWeight.setCellValue(elementRegion.dateWeightDouble());
//                    cellRegionsBrestWeekWeight.setCellValue(elementRegion.weekWeightDouble());
//                    cellBrest2021.setCellValue(regions.get(k).date2021Double());
//                    cellBrest2022.setCellValue(regions.get(k).date2022Double());
//                    break;
//                case 2:
//                    cellRegionsVitebskDateWeight.setCellValue(elementRegion.dateWeightDouble());
//                    cellRegionsVitebskWeekWeight.setCellValue(elementRegion.weekWeightDouble());
//                    cellVitebsk2021.setCellValue(regions.get(k).date2021Double());
//                    cellVitebsk2022.setCellValue(regions.get(k).date2022Double());
//                    break;
//                case 3:
//                    cellRegionsGomelDateWeight.setCellValue(elementRegion.dateWeightDouble());
//                    cellRegionsGomelWeekWeight.setCellValue(elementRegion.weekWeightDouble());
//                    cellGomel2021.setCellValue(regions.get(k).date2021Double());
//                    cellGomel2022.setCellValue(regions.get(k).date2022Double());
//                    break;
//                case 4:
//                    cellRegionsGrodnoDateWeight.setCellValue(elementRegion.dateWeightDouble());
//                    cellRegionsGrodnoWeekWeight.setCellValue(elementRegion.weekWeightDouble());
//                    cellGrodno2021.setCellValue(regions.get(k).date2021Double());
//                    cellGrodno2022.setCellValue(regions.get(k).date2022Double());
//                    break;
//                case 5:
//                    cellRegionsMinskDateWeight.setCellValue(elementRegion.dateWeightDouble());
//                    cellRegionsMinskWeekWeight.setCellValue(elementRegion.weekWeightDouble());
//                    cellMinsk2021.setCellValue(regions.get(k).date2021Double());
//                    cellMinsk2022.setCellValue(regions.get(k).date2022Double());
//                    break;
//                case 6:
//                    cellRegionsMogilevDateWeight.setCellValue(elementRegion.dateWeightDouble());
//                    cellRegionsMogilexWeekWeight.setCellValue(elementRegion.weekWeightDouble());
//                    cellMogilev2021.setCellValue(regions.get(k).date2021Double());
//                    cellMogilev2022.setCellValue(regions.get(k).date2022Double());
//                    break;
//
//            }
//        }
//    }

//    public static void createOneRows(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, ElementMass mass, int i) {
//
//        XSSFCell cellMassProductDateFromTo = row.createCell(i);
//        cellMassProductDateFromTo.setCellStyle(cellStyle);
//        cellMassProductDateFromTo.setCellValue(mass.dateFromToDouble());
//
//    }

//    /public static void createRows20212(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, List<ElementRegion> region, int i) {
//
//        XSSFCell cellBrest2021 = row.createCell(i);
//        cellBrest2021.setCellStyle(cellStyle);
//        cellBrest2021.setCellValue(region.get(0).date2021Double());
//
//        XSSFCell cellVitebsk2021 = row.createCell(i + 1);
//        cellVitebsk2021.setCellStyle(cellStyle);
//        cellVitebsk2021.setCellValue(region.get(1).date2021Double());
//
//        XSSFCell cellGomel2021 = row.createCell(i + 2);
//        cellGomel2021.setCellStyle(cellStyle);
//        cellGomel2021.setCellValue(region.get(2).date2021Double());
//
//        XSSFCell cellGrodno2021 = row.createCell(i + 3);
//        cellGrodno2021.setCellStyle(cellStyle);
//        cellGrodno2021.setCellValue(region.get(3).date2021Double());
//
//        XSSFCell cellMinsk2021 = row.createCell(i + 4);
//        cellMinsk2021.setCellStyle(cellStyle);
//        cellMinsk2021.setCellValue(region.get(4).date2021Double());
//
//        XSSFCell cellMogilev2021 = row.createCell(i + 5);
//        cellMogilev2021.setCellStyle(cellStyle);
//        cellMogilev2021.setCellValue(region.get(5).date2021Double());
//
//        XSSFCell cellBrest2022 = row.createCell(i + 6);
//        cellBrest2022.setCellStyle(cellStyle);
//        cellBrest2022.setCellValue(region.get(0).date2022Double());
//        XSSFCell cellVitebsk2022 = row.createCell(i + 7);
//        cellVitebsk2022.setCellStyle(cellStyle);
//        cellVitebsk2022.setCellValue(region.get(1).date2022Double());
//        XSSFCell cellGomel2022 = row.createCell(i + 8);
//        cellGomel2022.setCellStyle(cellStyle);
//        cellGomel2022.setCellValue(region.get(2).date2022Double());
//        XSSFCell cellGrodno2022 = row.createCell(i + 9);
//        cellGrodno2022.setCellStyle(cellStyle);
//        cellGrodno2022.setCellValue(region.get(3).date2022Double());
//        XSSFCell cellMinsk2022 = row.createCell(i + 10);
//        cellMinsk2022.setCellStyle(cellStyle);
//        cellMinsk2022.setCellValue(region.get(4).date2022Double());
//        XSSFCell cellMogilev2022 = row.createCell(i + 11);
//        cellMogilev2022.setCellStyle(cellStyle);
//        cellMogilev2022.setCellValue(region.get(5).date2022Double());
//
//
//    }

//    public static void createRowsFss2021(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, List<ElementRegion> regions, ElementMass mass, String country, int i) {
//
//        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
//        int rowLast = sheet.getLastRowNum();
//
//
//        XSSFRow row = sheet.createRow(rowLast + 1);
//        XSSFCell fss2021 = row.createCell(0);
//        XSSFCell fss22021 = row.createCell(1);
//        XSSFCell fss32021 = row.createCell(2);
//        fss2021.setCellStyle(cellStyle);
//        fss22021.setCellStyle(cellStyle);
//        fss32021.setCellStyle(cellStyle);
//        row.setHeight((short) 900);
//        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 2));
//        fss2021.setCellValue("Выдано ФСС, оформленных на реэкспорт (на всю продукцию), шт в 2021 г.");
//
//        XSSFCell fssPeriod2022 = row.createCell(i + 1);
//        fssPeriod2022.setCellStyle(cellStyle);
//        XSSFCell fssPeriod2021 = row.createCell(i + 2);
//        fssPeriod2021.setCellStyle(cellStyle);
//        XSSFCell fssPeriod7 = row.createCell(i + 3);
//        fssPeriod7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsBrest2021 = row.createCell(i + 4);
//        cellRegionsBrest2021.setCellStyle(cellStyle);
//        XSSFCell cellRegionsBrest2021_7 = row.createCell(i + 5);
//        cellRegionsBrest2021_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsVitebsk2021 = row.createCell(i + 6);
//        cellRegionsVitebsk2021.setCellStyle(cellStyle);
//        XSSFCell cellRegionsVitebsk2021_7 = row.createCell(i + 7);
//        cellRegionsVitebsk2021_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsGomel2021 = row.createCell(i + 8);
//        cellRegionsGomel2021.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGomel2021_7 = row.createCell(i + 9);
//        cellRegionsGomel2021_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsGrodno2021 = row.createCell(i + 10);
//        cellRegionsGrodno2021.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGrodno2021_7 = row.createCell(i + 11);
//        cellRegionsGrodno2021_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsMinsk2021 = row.createCell(i + 12);
//        cellRegionsMinsk2021.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMinsk2021_7 = row.createCell(i + 13);
//        cellRegionsMinsk2021_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsMogilev2021 = row.createCell(i + 14);
//        cellRegionsMogilev2021.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMogilev2021_7 = row.createCell(i + 15);
//        cellRegionsMogilev2021_7.setCellStyle(cellStyle);
//
////        fssPeriod2021.setCellValue(countryRequest.getFssPeriod2021_21());
////        fssPeriod2021.setCellValue(countryRequest.getFssPeriod2022_21());
////        fssPeriod7.setCellValue(countryRequest.getFssPeriod7_21());
////
////        cellRegionsBrest2021.setCellValue(countryRequest.getFss2021Brest());
////        cellRegionsVitebsk2021.setCellValue(countryRequest.getFss2021Vitebsk());
////        cellRegionsGomel2021.setCellValue(countryRequest.getFss2021Gomel());
////        cellRegionsGrodno2021.setCellValue(countryRequest.getFss2021Grodno());
////        cellRegionsMinsk2021.setCellValue(countryRequest.getFss2021Minsk());
////        cellRegionsMogilev2021.setCellValue(countryRequest.getFss2021Mogilev());
////
////        cellRegionsBrest2021_7.setCellValue(countryRequest.getFss2021Brest_7());
////        cellRegionsVitebsk2021_7.setCellValue(countryRequest.getFss2021Vitebsk_7());
////        cellRegionsGomel2021_7.setCellValue(countryRequest.getFss2021Gomel_7());
////        cellRegionsGrodno2021_7.setCellValue(countryRequest.getFss2021Grodno_7());
////        cellRegionsMinsk2021_7.setCellValue(countryRequest.getFss2021Minsk_7());
////        cellRegionsMogilev2021_7.setCellValue(countryRequest.getFss2021Mogilev_7());
//        // }
//
//
//    }

//    public static void createRowsFss2022(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, List<ElementRegion> regions, ElementMass mass, String country, int i) {
//
//        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
//        int rowLast = sheet.getLastRowNum();
//
//
//        XSSFRow row = sheet.createRow(rowLast + 1);
//        XSSFCell fss2022 = row.createCell(0);
//        XSSFCell fss22022 = row.createCell(1);
//        XSSFCell fss32022 = row.createCell(2);
//        fss2022.setCellStyle(cellStyle);
//        fss22022.setCellStyle(cellStyle);
//        fss32022.setCellStyle(cellStyle);
//        row.setHeight((short) 900);
//        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 2));
//        fss2022.setCellValue("Выдано ФСС, оформленных на реэкспорт             (на всю продукцию), шт в 2022 г.");
//
//        XSSFCell fssPeriod2022 = row.createCell(i + 1);
//        fssPeriod2022.setCellStyle(cellStyle);
//        XSSFCell fssPeriod2021 = row.createCell(i + 2);
//        fssPeriod2021.setCellStyle(cellStyle);
//        XSSFCell fssPeriod7 = row.createCell(i + 3);
//        fssPeriod7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsBrest2022 = row.createCell(i + 4);
//        cellRegionsBrest2022.setCellStyle(cellStyle);
//        XSSFCell cellRegionsBrest2022_7 = row.createCell(i + 5);
//        cellRegionsBrest2022_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsVitebsk2022 = row.createCell(i + 6);
//        cellRegionsVitebsk2022.setCellStyle(cellStyle);
//        XSSFCell cellRegionsVitebsk2022_7 = row.createCell(i + 7);
//        cellRegionsVitebsk2022_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsGomel2022 = row.createCell(i + 8);
//        cellRegionsGomel2022.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGomel2022_7 = row.createCell(i + 9);
//        cellRegionsGomel2022_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsGrodno2022 = row.createCell(i + 10);
//        cellRegionsGrodno2022.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGrodno2022_7 = row.createCell(i + 11);
//        cellRegionsGrodno2022_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsMinsk2022 = row.createCell(i + 12);
//        cellRegionsMinsk2022.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMinsk2022_7 = row.createCell(i + 13);
//        cellRegionsMinsk2022_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsMogilev2022 = row.createCell(i + 14);
//        cellRegionsMogilev2022.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMogilev2022_7 = row.createCell(i + 15);
//        cellRegionsMogilev2022_7.setCellStyle(cellStyle);
//
////        fssPeriod2022.setCellValue(countryRequest.getFssPeriod2021_22());
////        fssPeriod2022.setCellValue(countryRequest.getFssPeriod2022_22());
////        fssPeriod7.setCellValue(countryRequest.getFssPeriod7_22());
//
////        cellRegionsBrest2022.setCellValue(countryRequest.getFss2022Brest());
////        cellRegionsVitebsk2022.setCellValue(countryRequest.getFss2022Vitebsk());
////        cellRegionsGomel2022.setCellValue(countryRequest.getFss2022Gomel());
////        cellRegionsGrodno2022.setCellValue(countryRequest.getFss2022Grodno());
////        cellRegionsMinsk2022.setCellValue(countryRequest.getFss2022Minsk());
////        cellRegionsMogilev2022.setCellValue(countryRequest.getFss2022Mogilev());
////
////        cellRegionsBrest2022_7.setCellValue(countryRequest.getFss2022Brest_7());
////        cellRegionsVitebsk2022_7.setCellValue(countryRequest.getFss2022Vitebsk_7());
////        cellRegionsGomel2022_7.setCellValue(countryRequest.getFss2022Gomel_7());
////        cellRegionsGrodno2022_7.setCellValue(countryRequest.getFss2022Grodno_7());
////        cellRegionsMinsk2022_7.setCellValue(countryRequest.getFss2022Minsk_7());
////        cellRegionsMogilev2022_7.setCellValue(countryRequest.getFss2022Mogilev_7());
//        // }
//
//
//    }


//    public static void createRowsAllFss(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, List<ElementRegion> regions, ElementMass mass, String country, int i) {
//
//        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
//        int rowLast = sheet.getLastRowNum();
//
//
//        XSSFRow row = sheet.createRow(rowLast + 1);
//        XSSFCell allfss = row.createCell(0);
//        XSSFCell allfss2 = row.createCell(1);
//        XSSFCell allfss3 = row.createCell(2);
//        allfss.setCellStyle(cellStyle);
//        allfss2.setCellStyle(cellStyle);
//        allfss3.setCellStyle(cellStyle);
//        row.setHeight((short) 900);
//        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 2));
//        allfss.setCellValue("Выдано ФСС, оформленных на реэкспорт (на всю продукцию), шт");
//
////        XSSFRow rowTotal4 = sheet.createRow(rowLast + 4);
////        XSSFCell fss = rowTotal4.createCell(0);
////        XSSFCell fss2 = rowTotal4.createCell(1);
////        XSSFCell fss3 = rowTotal4.createCell(2);
////        fss.setCellStyle(cellStyle);
////        fss2.setCellStyle(cellStyle);
////        fss3.setCellStyle(cellStyle);
////        rowTotal4.setHeight((short) 900);
////        sheet.addMergedRegion(new CellRangeAddress(rowLast + 4, rowLast + 4, 0, 2));
//
//
//        XSSFCell fssPeriod2022 = row.createCell(i + 1);
//        fssPeriod2022.setCellStyle(cellStyle);
//        XSSFCell fssPeriod2021 = row.createCell(i + 2);
//        fssPeriod2021.setCellStyle(cellStyle);
//        XSSFCell fssPeriod7 = row.createCell(i + 3);
//        fssPeriod7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsBrest = row.createCell(i + 4);
//        cellRegionsBrest.setCellStyle(cellStyle);
//        XSSFCell cellRegionsBrest_7 = row.createCell(i + 5);
//        cellRegionsBrest_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsVitebsk = row.createCell(i + 6);
//        cellRegionsVitebsk.setCellStyle(cellStyle);
//        XSSFCell cellRegionsVitebsk_7 = row.createCell(i + 7);
//        cellRegionsVitebsk_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsGomel = row.createCell(i + 8);
//        cellRegionsGomel.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGomel_7 = row.createCell(i + 9);
//        cellRegionsGomel_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsGrodno = row.createCell(i + 10);
//        cellRegionsGrodno.setCellStyle(cellStyle);
//        XSSFCell cellRegionsGrodno_7 = row.createCell(i + 11);
//        cellRegionsGrodno_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsMinsk = row.createCell(i + 12);
//        cellRegionsMinsk.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMinsk_7 = row.createCell(i + 13);
//        cellRegionsMinsk_7.setCellStyle(cellStyle);
//
//        XSSFCell cellRegionsMogilev = row.createCell(i + 14);
//        cellRegionsMogilev.setCellStyle(cellStyle);
//        XSSFCell cellRegionsMogilev_7 = row.createCell(i + 15);
//        cellRegionsMogilev_7.setCellStyle(cellStyle);
//
//        fssPeriod2021.setCellValue(" *** ");
//        fssPeriod2022.setCellValue(countryRequest.getAllFss());
//        fssPeriod7.setCellValue(countryRequest.getAllFss_7());
//
//        cellRegionsBrest.setCellValue(countryRequest.getAllFssBrest());
//        cellRegionsVitebsk.setCellValue(countryRequest.getAllFssVitebsk());
//        cellRegionsGomel.setCellValue(countryRequest.getAllFssGomel());
//        cellRegionsGrodno.setCellValue(countryRequest.getAllFssGrodno());
//        cellRegionsMinsk.setCellValue(countryRequest.getAllFssMinsk());
//        cellRegionsMogilev.setCellValue(countryRequest.getAllFssMogilev());
//
//        cellRegionsBrest_7.setCellValue(countryRequest.getAllFss_7Brest());
//        cellRegionsVitebsk_7.setCellValue(countryRequest.getAllFss_7Vitebsk());
//        cellRegionsGomel_7.setCellValue(countryRequest.getAllFss_7Gomel());
//        cellRegionsGrodno_7.setCellValue(countryRequest.getAllFss_7Grodno());
//        cellRegionsMinsk_7.setCellValue(countryRequest.getAllFss_7Minsk());
//        cellRegionsMogilev_7.setCellValue(countryRequest.getAllFss_7Mogilev());
//        // }
//
//
//    }

    public static void createRowsMaterial(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, int i) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast = sheet.getLastRowNum();

        XSSFRow row = sheet.createRow(rowLast + 1);
        XSSFCell cellCountry = row.createCell(0);
        XSSFCell cellCountry2 = row.createCell(1);

        cellCountry.setCellStyle(cellStyle);
        cellCountry2.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 1));

        XSSFRow row2 = sheet.createRow(rowLast + 2);
        XSSFCell cellCountry3 = row2.createCell(0);
        XSSFCell cellCountry33 = row2.createCell(1);

        cellCountry3.setCellStyle(cellStyle);
        cellCountry33.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 2, rowLast + 2, 0, 1));

        cellCountry3.setCellValue("Срезы и посадочный материал цветочной и лесодекоративной, горшечной продукции");
        row2.setHeight((short) 1200);

        XSSFCell massa = row.createCell(2);
        massa.setCellStyle(cellStyle);
        massa.setCellValue("тр. ед.");

        XSSFCell massa2 = row2.createCell(2);
        massa2.setCellStyle(cellStyle);
        massa2.setCellValue("млн. шт.");


        XSSFCell unit_materialFromTo = row.createCell(i + 1);
        unit_materialFromTo.setCellStyle(cellStyle);
        XSSFCell unit_material_7 = row.createCell(i + 2);
        unit_material_7.setCellStyle(cellStyle);
        XSSFCell unit_material_2022 = row.createCell(i + 3);
        unit_material_2022.setCellStyle(cellStyle);

        XSSFCell unit_materialBrest = row.createCell(i + 4);
        unit_materialBrest.setCellStyle(cellStyle);
        XSSFCell unit_material_7Brest = row.createCell(i + 5);
        unit_material_7Brest.setCellStyle(cellStyle);

        XSSFCell unit_materialVitebsk = row.createCell(i + 6);
        unit_materialVitebsk.setCellStyle(cellStyle);
        XSSFCell unit_material_7Vitebsk = row.createCell(i + 7);
        unit_material_7Vitebsk.setCellStyle(cellStyle);

        XSSFCell unit_materialGomel = row.createCell(i + 8);
        unit_materialGomel.setCellStyle(cellStyle);
        XSSFCell unit_material_7Gomel = row.createCell(i + 9);
        unit_material_7Gomel.setCellStyle(cellStyle);

        XSSFCell unit_materialGrodno = row.createCell(i + 10);
        unit_materialGrodno.setCellStyle(cellStyle);
        XSSFCell unit_material_7Grodno = row.createCell(i + 11);
        unit_material_7Grodno.setCellStyle(cellStyle);

        XSSFCell unit_materialMinsk = row.createCell(i + 12);
        unit_materialMinsk.setCellStyle(cellStyle);
        XSSFCell unit_materials_7Minsk = row.createCell(i + 13);
        unit_materials_7Minsk.setCellStyle(cellStyle);

        XSSFCell unit_materialMogilev = row.createCell(i + 14);
        unit_materialMogilev.setCellStyle(cellStyle);
        XSSFCell unit_material_7Mogilev = row.createCell(i + 15);
        unit_material_7Mogilev.setCellStyle(cellStyle);

        unit_materialFromTo.setCellValue(countryRequest.getUnit_materialFromTo());
        unit_material_7.setCellValue(countryRequest.getUnit_material_7());
        unit_material_2022.setCellValue(countryRequest.getUnit_material_2022());

        unit_materialBrest.setCellValue(countryRequest.getUnit_materialBrest());
        unit_material_7Brest.setCellValue(countryRequest.getUnit_material_7Brest());
        unit_materialVitebsk.setCellValue(countryRequest.getUnit_materialVitebsk());
        unit_material_7Vitebsk.setCellValue(countryRequest.getUnit_material_7Vitebsk());
        unit_materialGomel.setCellValue(countryRequest.getUnit_materialGomel());
        unit_material_7Gomel.setCellValue(countryRequest.getUnit_material_7Gomel());
        unit_materialGrodno.setCellValue(countryRequest.getUnit_materialGrodno());
        unit_material_7Grodno.setCellValue(countryRequest.getUnit_material_7Grodno());
        unit_materialMinsk.setCellValue(countryRequest.getUnit_materialMinsk());
        unit_materials_7Minsk.setCellValue(countryRequest.getUnit_materials_7Minsk());
        unit_materialMogilev.setCellValue(countryRequest.getUnit_materialMogilev());
        unit_material_7Mogilev.setCellValue(countryRequest.getUnit_material_7Mogilev());

        XSSFCell piece_material2022 = row2.createCell(i + 1);
        piece_material2022.setCellStyle(cellStyle);
        XSSFCell piece_material_7 = row2.createCell(i + 2);
        piece_material_7.setCellStyle(cellStyle);
        XSSFCell piece_material2021 = row2.createCell(i + 3);
        piece_material2021.setCellStyle(cellStyle);


        XSSFCell piece_materialBrest = row2.createCell(i + 4);
        piece_materialBrest.setCellStyle(cellStyle);
        XSSFCell piece_material_7Brest = row2.createCell(i + 5);
        piece_material_7Brest.setCellStyle(cellStyle);

        XSSFCell piece_materialVitebsk = row2.createCell(i + 6);
        piece_materialVitebsk.setCellStyle(cellStyle);
        XSSFCell piece_material_7Vitebsk = row2.createCell(i + 7);
        piece_material_7Vitebsk.setCellStyle(cellStyle);

        XSSFCell piece_materialGomel = row2.createCell(i + 8);
        piece_materialGomel.setCellStyle(cellStyle);
        XSSFCell piece_material_7Gomel = row2.createCell(i + 9);
        piece_material_7Gomel.setCellStyle(cellStyle);

        XSSFCell piece_materialGrodno = row2.createCell(i + 10);
        piece_materialGrodno.setCellStyle(cellStyle);
        XSSFCell piece_material_7Grodno = row2.createCell(i + 11);
        piece_material_7Grodno.setCellStyle(cellStyle);

        XSSFCell piece_materialMinsk = row2.createCell(i + 12);
        piece_materialMinsk.setCellStyle(cellStyle);
        XSSFCell piece_materials_7Minsk = row2.createCell(i + 13);
        piece_materials_7Minsk.setCellStyle(cellStyle);

        XSSFCell piece_materialMogilev = row2.createCell(i + 14);
        piece_materialMogilev.setCellStyle(cellStyle);
        XSSFCell piece_material_7Mogilev = row2.createCell(i + 15);
        piece_material_7Mogilev.setCellStyle(cellStyle);

        piece_material2021.setCellValue(" *** ");
        piece_material2022.setCellValue(countryRequest.getPiece_material());
        piece_material_7.setCellValue(countryRequest.getPiece_material_7());

        piece_materialBrest.setCellValue(countryRequest.getPiece_materialBrest());
        piece_material_7Brest.setCellValue(countryRequest.getPiece_material_7Brest());
        piece_materialVitebsk.setCellValue(countryRequest.getPiece_materialVitebsk());
        piece_material_7Vitebsk.setCellValue(countryRequest.getPiece_material_7Vitebsk());
        piece_materialGomel.setCellValue(countryRequest.getPiece_materialGomel());
        piece_material_7Gomel.setCellValue(countryRequest.getPiece_material_7Gomel());
        piece_materialGrodno.setCellValue(countryRequest.getPiece_materialGrodno());
        piece_material_7Grodno.setCellValue(countryRequest.getPiece_material_7Grodno());
        piece_materialMinsk.setCellValue(countryRequest.getPiece_materialMinsk());
        piece_materials_7Minsk.setCellValue(countryRequest.getPiece_materials_7Minsk());
        piece_materialMogilev.setCellValue(countryRequest.getPiece_materialMogilev());
        piece_material_7Mogilev.setCellValue(countryRequest.getPiece_material_7Mogilev());

    }
}