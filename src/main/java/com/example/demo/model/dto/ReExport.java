package com.example.demo.model.dto;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class ReExport {

    public static void createRowsMaterial(XSSFWorkbook xssfWorkbook, CountryReport countryRequest, XSSFCellStyle cellStyle, XSSFCellStyle cellStyleRow, int i) {

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

        XSSFCell mass = row.createCell(2);
        mass.setCellStyle(cellStyle);
        mass.setCellValue("тр. ед.");

        XSSFCell mass2 = row2.createCell(2);
        mass2.setCellStyle(cellStyle);
        mass2.setCellValue("млн. шт.");


        XSSFCell unit_materialFromTo = row.createCell(i + 1);
        unit_materialFromTo.setCellStyle(cellStyleRow);
        XSSFCell unit_material_7 = row.createCell(i + 2);
        unit_material_7.setCellStyle(cellStyleRow);
        XSSFCell unit_material_2022 = row.createCell(i + 3);
        unit_material_2022.setCellStyle(cellStyleRow);

        XSSFCell unit_materialBrest = row.createCell(i + 4);
        unit_materialBrest.setCellStyle(cellStyleRow);
        XSSFCell unit_material_7Brest = row.createCell(i + 5);
        unit_material_7Brest.setCellStyle(cellStyleRow);

        XSSFCell unit_materialVitebsk = row.createCell(i + 6);
        unit_materialVitebsk.setCellStyle(cellStyleRow);
        XSSFCell unit_material_7Vitebsk = row.createCell(i + 7);
        unit_material_7Vitebsk.setCellStyle(cellStyleRow);

        XSSFCell unit_materialGomel = row.createCell(i + 8);
        unit_materialGomel.setCellStyle(cellStyleRow);
        XSSFCell unit_material_7Gomel = row.createCell(i + 9);
        unit_material_7Gomel.setCellStyle(cellStyleRow);

        XSSFCell unit_materialGrodno = row.createCell(i + 10);
        unit_materialGrodno.setCellStyle(cellStyleRow);
        XSSFCell unit_material_7Grodno = row.createCell(i + 11);
        unit_material_7Grodno.setCellStyle(cellStyleRow);

        XSSFCell unit_materialMinsk = row.createCell(i + 12);
        unit_materialMinsk.setCellStyle(cellStyleRow);
        XSSFCell unit_materials_7Minsk = row.createCell(i + 13);
        unit_materials_7Minsk.setCellStyle(cellStyleRow);

        XSSFCell unit_materialMogilev = row.createCell(i + 14);
        unit_materialMogilev.setCellStyle(cellStyleRow);
        XSSFCell unit_material_7Mogilev = row.createCell(i + 15);
        unit_material_7Mogilev.setCellStyle(cellStyleRow);

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

        XSSFCell piece_material2022 = row2.createCell(i + 1);
        piece_material2022.setCellStyle(cellStyleRow);
        XSSFCell piece_material_7 = row2.createCell(i + 2);
        piece_material_7.setCellStyle(cellStyleRow);
        XSSFCell piece_material2021 = row2.createCell(i + 3);
        piece_material2021.setCellStyle(cellStyleRow);


        XSSFCell piece_materialBrest = row2.createCell(i + 4);
        piece_materialBrest.setCellStyle(cellStyleRow);
        XSSFCell piece_material_7Brest = row2.createCell(i + 5);
        piece_material_7Brest.setCellStyle(cellStyleRow);

        XSSFCell piece_materialVitebsk = row2.createCell(i + 6);
        piece_materialVitebsk.setCellStyle(cellStyleRow);
        XSSFCell piece_material_7Vitebsk = row2.createCell(i + 7);
        piece_material_7Vitebsk.setCellStyle(cellStyleRow);

        XSSFCell piece_materialGomel = row2.createCell(i + 8);
        piece_materialGomel.setCellStyle(cellStyleRow);
        XSSFCell piece_material_7Gomel = row2.createCell(i + 9);
        piece_material_7Gomel.setCellStyle(cellStyleRow);

        XSSFCell piece_materialGrodno = row2.createCell(i + 10);
        piece_materialGrodno.setCellStyle(cellStyleRow);
        XSSFCell piece_material_7Grodno = row2.createCell(i + 11);
        piece_material_7Grodno.setCellStyle(cellStyleRow);

        XSSFCell piece_materialMinsk = row2.createCell(i + 12);
        piece_materialMinsk.setCellStyle(cellStyleRow);
        XSSFCell piece_materials_7Minsk = row2.createCell(i + 13);
        piece_materials_7Minsk.setCellStyle(cellStyleRow);

        XSSFCell piece_materialMogilev = row2.createCell(i + 14);
        piece_materialMogilev.setCellStyle(cellStyleRow);
        XSSFCell piece_material_7Mogilev = row2.createCell(i + 15);
        piece_material_7Mogilev.setCellStyle(cellStyleRow);

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
}