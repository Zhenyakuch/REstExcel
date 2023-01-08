package com.example.demo.model.dto;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public class Import {

    public static void createRowsImport(XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, List<ElementRegion> regions, ElementMass mass, String country, int i) {

//        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
//        cellStyle.setWrapText(true);
//        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
//        cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
//        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);


        XSSFCell cellCountry = row.createCell(1);
        cellCountry.setCellStyle(cellStyle);
        XSSFCell cellMassProductDateWeight = row.createCell(2);
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

        cellCountry.setCellValue(country);
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
}
