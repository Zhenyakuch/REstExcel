package com.example.demo.model.dto;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class ReExport {

    public static void createRowsMaterial(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow, int i) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast = sheet.getLastRowNum();
        XSSFRow row = sheet.createRow(rowLast + 1);

        XSSFCell cellCountry = getXssfCell(row, 0, cellStyle);
        XSSFCell cellCountry2 = getXssfCell(row, 1, cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 1));

        XSSFRow row2 = sheet.createRow(rowLast + 2);
        XSSFCell cellCountry3 = getXssfCell(row2, 0, cellStyle);
        XSSFCell cellCountry33 = getXssfCell(row2, 1, cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 2, rowLast + 2, 0, 1));

        cellCountry3.setCellValue("Срезы и посадочный материал цветочной и лесодекоративной, горшечной продукции");
        row2.setHeight((short) 1200);

        XSSFCell mass = getXssfCell(row, 2, cellStyle);
        mass.setCellValue("тр. ед.");

        XSSFCell mass2 = getXssfCell(row2, 2, cellStyle);
        mass2.setCellValue("млн. шт.");


        XSSFCell unit_materialFromTo = getXssfCell(row, i + 1, cellStyleRow);
        XSSFCell unit_material_7 = getXssfCell(row, i + 2, cellStyleRow);
        XSSFCell unit_material_2022 = getXssfCell(row, i + 3, cellStyleRow);
        XSSFCell unit_materialBrest = getXssfCell(row, i + 4, cellStyleRow);
        XSSFCell unit_material_7Brest = getXssfCell(row, i + 5, cellStyleRow);
        XSSFCell unit_materialVitebsk = getXssfCell(row, i + 6, cellStyleRow);
        XSSFCell unit_material_7Vitebsk = getXssfCell(row, i + 7, cellStyleRow);
        XSSFCell unit_materialGomel = getXssfCell(row, i + 8, cellStyleRow);
        XSSFCell unit_material_7Gomel = getXssfCell(row, i + 9, cellStyleRow);
        XSSFCell unit_materialGrodno = getXssfCell(row, i + 10, cellStyleRow);
        XSSFCell unit_material_7Grodno = getXssfCell(row, i + 11, cellStyleRow);
        XSSFCell unit_materialMinsk = getXssfCell(row, i + 12, cellStyleRow);
        XSSFCell unit_materials_7Minsk = getXssfCell(row, i + 13, cellStyleRow);
        XSSFCell unit_materialMogilev = getXssfCell(row, i + 14, cellStyleRow);
        XSSFCell unit_material_7Mogilev = getXssfCell(row, i + 15, cellStyleRow);

        unit_materialFromTo.setCellValue(countryRequest.getUnits().get(0).getUnit_materialFromTo());
        unit_material_7.setCellValue(countryRequest.getUnits().get(0).getUnit_material_7());
        unit_material_2022.setCellValue(countryRequest.getUnits().get(0).getUnit_material_2022());

        unit_materialBrest.setCellValue(countryRequest.getUnits().get(0).getUnit_materialBrest());
        unit_material_7Brest.setCellValue(countryRequest.getUnits().get(0).getUnit_material_7Brest());
        unit_materialVitebsk.setCellValue(countryRequest.getUnits().get(0).getUnit_materialVitebsk());
        unit_material_7Vitebsk.setCellValue(countryRequest.getUnits().get(0).getUnit_material_7Vitebsk());
        unit_materialGomel.setCellValue(countryRequest.getUnits().get(0).getUnit_materialGomel());
        unit_material_7Gomel.setCellValue(countryRequest.getUnits().get(0).getUnit_material_7Gomel());
        unit_materialGrodno.setCellValue(countryRequest.getUnits().get(0).getUnit_materialGrodno());
        unit_material_7Grodno.setCellValue(countryRequest.getUnits().get(0).getUnit_material_7Grodno());
        unit_materialMinsk.setCellValue(countryRequest.getUnits().get(0).getUnit_materialMinsk());
        unit_materials_7Minsk.setCellValue(countryRequest.getUnits().get(0).getUnit_materials_7Minsk());
        unit_materialMogilev.setCellValue(countryRequest.getUnits().get(0).getUnit_materialMogilev());
        unit_material_7Mogilev.setCellValue(countryRequest.getUnits().get(0).getUnit_material_7Mogilev());

        XSSFCell piece_material2022 = getXssfCell(row2, i + 1, cellStyleRow);
        XSSFCell piece_material_7 = getXssfCell(row2, i + 2, cellStyleRow);
        XSSFCell piece_material2021 = getXssfCell(row2, i + 3, cellStyleRow);
        XSSFCell piece_materialBrest = getXssfCell(row2, i + 4, cellStyleRow);
        XSSFCell piece_material_7Brest = getXssfCell(row2, i + 5, cellStyleRow);
        XSSFCell piece_materialVitebsk = getXssfCell(row2, i + 6, cellStyleRow);
        XSSFCell piece_material_7Vitebsk = getXssfCell(row2, i + 7, cellStyleRow);
        XSSFCell piece_materialGomel = getXssfCell(row2, i + 8, cellStyleRow);
        XSSFCell piece_material_7Gomel = getXssfCell(row2, i + 9, cellStyleRow);
        XSSFCell piece_materialGrodno = getXssfCell(row2, i + 10, cellStyleRow);
        XSSFCell piece_material_7Grodno = getXssfCell(row2, i + 11, cellStyleRow);
        XSSFCell piece_materialMinsk = getXssfCell(row2, i + 12, cellStyleRow);
        XSSFCell piece_materials_7Minsk = getXssfCell(row2, i + 13, cellStyleRow);
        XSSFCell piece_materialMogilev = getXssfCell(row2, i + 14, cellStyleRow);
        XSSFCell piece_material_7Mogilev = getXssfCell(row2, i + 15, cellStyleRow);

        piece_material2021.setCellValue(" *** ");
        piece_material2022.setCellValue(countryRequest.getPiece().get(0).getPiece_material());
        piece_material_7.setCellValue(countryRequest.getPiece().get(0).getPiece_material_7());

        piece_materialBrest.setCellValue(countryRequest.getPiece().get(0).getPiece_materialBrest());
        piece_material_7Brest.setCellValue(countryRequest.getPiece().get(0).getPiece_material_7Brest());
        piece_materialVitebsk.setCellValue(countryRequest.getPiece().get(0).getPiece_materialVitebsk());
        piece_material_7Vitebsk.setCellValue(countryRequest.getPiece().get(0).getPiece_material_7Vitebsk());
        piece_materialGomel.setCellValue(countryRequest.getPiece().get(0).getPiece_materialGomel());
        piece_material_7Gomel.setCellValue(countryRequest.getPiece().get(0).getPiece_material_7Gomel());
        piece_materialGrodno.setCellValue(countryRequest.getPiece().get(0).getPiece_materialGrodno());
        piece_material_7Grodno.setCellValue(countryRequest.getPiece().get(0).getPiece_material_7Grodno());
        piece_materialMinsk.setCellValue(countryRequest.getPiece().get(0).getPiece_materialMinsk());
        piece_materials_7Minsk.setCellValue(countryRequest.getPiece().get(0).getPiece_materials_7Minsk());
        piece_materialMogilev.setCellValue(countryRequest.getPiece().get(0).getPiece_materialMogilev());
        piece_material_7Mogilev.setCellValue(countryRequest.getPiece().get(0).getPiece_material_7Mogilev());

    }

    private static XSSFCell getXssfCell(XSSFRow row, int i, XSSFCellStyle cellStyleRow) {
        XSSFCell cellRegionsBrestDateWeight = row.createCell(i);
        cellRegionsBrestDateWeight.setCellStyle(cellStyleRow);
        return cellRegionsBrestDateWeight;
    }

}