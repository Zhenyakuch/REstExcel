package com.example.demo.model.dto;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class ReExport {

    public static void createRowsReExport(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, List<ElementRegion> regions, ElementMass mass, String country, int i) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast33 = sheet.getLastRowNum();
        XSSFCell cellCountry = row.createCell(1);
        XSSFCell cellCountry2 = row.createCell(2);
        cellCountry.setCellStyle(cellStyle);
        cellCountry2.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast33, rowLast33, 1, 2));
        cellCountry.setCellValue(country);

        XSSFCell cellMassProductDateWeight = row.createCell(3);
        cellMassProductDateWeight.setCellStyle(cellStyle);
        XSSFCell cellMassProductWeekWeight = row.createCell(i);
        cellMassProductWeekWeight.setCellStyle(cellStyle);
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

        cellMassProductDateWeight.setCellValue(mass.dateWeightDouble());
        cellMassProductWeekWeight.setCellValue(mass.weekWeightDouble());

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

    public static void createOneRows(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, ElementMass mass, int i) {

        XSSFCell cellMassProductDateFromTo = row.createCell(i);
        cellMassProductDateFromTo.setCellStyle(cellStyle);
        cellMassProductDateFromTo.setCellValue(mass.dateFromToDouble());

    }

    public static void createRows20212(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, List<ElementRegion> region, int i) {

        XSSFCell cellBrest2021 = row.createCell(i);
        cellBrest2021.setCellStyle(cellStyle);
        cellBrest2021.setCellValue(region.get(0).getDate2021());
        XSSFCell cellVitebsk2021 = row.createCell(i + 1);
        cellVitebsk2021.setCellStyle(cellStyle);
        cellVitebsk2021.setCellValue(region.get(1).getDate2021());
        XSSFCell cellGomel2021 = row.createCell(i + 2);
        cellGomel2021.setCellStyle(cellStyle);
        cellGomel2021.setCellValue(region.get(2).getDate2021());
        XSSFCell cellGrodno2021 = row.createCell(i + 3);
        cellGrodno2021.setCellStyle(cellStyle);
        cellGrodno2021.setCellValue(region.get(3).getDate2021());
        XSSFCell cellMinsk2021 = row.createCell(i + 4);
        cellMinsk2021.setCellStyle(cellStyle);
        cellMinsk2021.setCellValue(region.get(4).getDate2021());
        XSSFCell cellMogilev2021 = row.createCell(i + 5);
        cellMogilev2021.setCellStyle(cellStyle);
        cellMogilev2021.setCellValue(region.get(5).getDate2021());

        XSSFCell cellBrest2022 = row.createCell(i + 6);
        cellBrest2022.setCellStyle(cellStyle);
        cellBrest2022.setCellValue(region.get(0).getDate2022());
        XSSFCell cellVitebsk2022 = row.createCell(i + 7);
        cellVitebsk2022.setCellStyle(cellStyle);
        cellVitebsk2022.setCellValue(region.get(1).getDate2022());
        XSSFCell cellGomel2022 = row.createCell(i + 8);
        cellGomel2022.setCellStyle(cellStyle);
        cellGomel2022.setCellValue(region.get(2).getDate2022());
        XSSFCell cellGrodno2022 = row.createCell(i + 9);
        cellGrodno2022.setCellStyle(cellStyle);
        cellGrodno2022.setCellValue(region.get(3).getDate2022());
        XSSFCell cellMinsk2022 = row.createCell(i + 10);
        cellMinsk2022.setCellStyle(cellStyle);
        cellMinsk2022.setCellValue(region.get(4).getDate2022());
        XSSFCell cellMogilev2022 = row.createCell(i + 11);
        cellMogilev2022.setCellStyle(cellStyle);
        cellMogilev2022.setCellValue(region.get(5).getDate2022());

    }
}
