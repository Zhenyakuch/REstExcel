package com.example.demo.model.dto;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class Tranzit {

    public static void createRows(int cellCount, XSSFWorkbook xssfWorkbook, XSSFRow row, XSSFCellStyle cellStyle, List<ElementRegion> regions, List<NamePoints> namePoints) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int rowLast2 = sheet.getLastRowNum();

        XSSFCell cellName_obl = row.createCell(0);
        cellName_obl.setCellStyle(cellStyle);
        XSSFCell cellName_obl1 = row.createCell(1);
        cellName_obl1.setCellStyle(cellStyle);
        XSSFCell cellName_obl2 = row.createCell(2);
        cellName_obl2.setCellStyle(cellStyle);
        XSSFCell cellName_obl3 = row.createCell(3);
        cellName_obl3.setCellStyle(cellStyle);
        XSSFCell cellName_obl4 = row.createCell(4);
        cellName_obl4.setCellStyle(cellStyle);
        XSSFCell cellName_obl5 = row.createCell(5);
        cellName_obl5.setCellStyle(cellStyle);
        XSSFCell cellName_obl6 = row.createCell(6);
        cellName_obl6.setCellStyle(cellStyle);
        XSSFCell cellName_obl7 = row.createCell(7);
        cellName_obl7.setCellStyle(cellStyle);
        XSSFCell cellName_obl8 = row.createCell(8);
        cellName_obl8.setCellStyle(cellStyle);
        XSSFCell cellName_obl9 = row.createCell(9);
        cellName_obl9.setCellStyle(cellStyle);
        XSSFCell cellName_obl10 = row.createCell(10);
        cellName_obl10.setCellStyle(cellStyle);
        XSSFCell cellName_obl11 = row.createCell(11);
        cellName_obl11.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowLast2, rowLast2, 0, 11));

       // for (int k = cellCount; regions.size() > k; k++) {
//            for (int k = 0; namePoints.size() > k; k++) {
            ElementRegion elementRegion = regions.get(cellCount);
            switch (elementRegion.getRegion()) {
                case 1:
                    cellName_obl.setCellValue("БРЕСТСКАЯ ОБЛАСТЬ");
                    for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                        cicle( sheet,namePoints,  cellCount, cellStyle);
                   }
                    break;
                case 2:
                    cellName_obl.setCellValue("ВИТЕБСКАЯ ОБЛАСТЬ");
                    for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                        cicle( sheet,namePoints,  cellCount, cellStyle);
                    }
                    break;
                case 3:
                    cellName_obl.setCellValue("ГОМЕЛЬСКАЯ ОБЛАСТЬ");
                    for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                        cicle( sheet,namePoints,  cellCount, cellStyle);
                    }
                    break;

                case 4:
                    cellName_obl.setCellValue("ГРОДНЕНСКАЯ ОБЛАСТЬ");
                    for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                        cicle( sheet,namePoints,  cellCount, cellStyle);
                    }
                    break;
                case 5:
                    cellName_obl.setCellValue("МИНСКАЯ ОБЛАСТЬ");
                    for (cellCount = 0; namePoints.size() > cellCount; cellCount++) {
                        cicle( sheet,namePoints,  cellCount, cellStyle);
                    }
                    break;

//            }
        }
    }

    public static void cicle( XSSFSheet sheet,List<NamePoints> namePoints, int k,XSSFCellStyle cellStyle){
//        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
//        int rowLast2 = sheet.getLastRowNum();
        int rowLast = sheet.getLastRowNum();
        XSSFRow row2 = sheet.createRow(rowLast + 1);
        XSSFCell cellName_points = row2.createCell(0);
        cellName_points.setCellStyle(cellStyle);
        XSSFCell cellThousand_t = row2.createCell(1);
        cellThousand_t.setCellStyle(cellStyle);
        XSSFCell cellThousand_pos_ed = row2.createCell(2);
        cellThousand_pos_ed.setCellStyle(cellStyle);
        XSSFCell cellThousand_sht = row2.createCell(3);
        cellThousand_sht.setCellStyle(cellStyle);
        XSSFCell cellThousand_part = row2.createCell(4);
        cellThousand_part.setCellStyle(cellStyle);
        XSSFCell cellThousand_m2 = row2.createCell(5);
        cellThousand_m2.setCellStyle(cellStyle);
        XSSFCell cellThousand_m3 = row2.createCell(6);
        cellThousand_m3.setCellStyle(cellStyle);
        XSSFCell cellRailway_wagons = row2.createCell(7);
        cellRailway_wagons.setCellStyle(cellStyle);
        XSSFCell cellMotor_transport = row2.createCell(8);
        cellMotor_transport.setCellStyle(cellStyle);
        XSSFCell cellContainer = row2.createCell(9);
        cellContainer.setCellStyle(cellStyle);
        XSSFCell cellBaggage = row2.createCell(10);
        cellBaggage.setCellStyle(cellStyle);
        XSSFCell cellAirplane = row2.createCell(11);
        cellAirplane.setCellStyle(cellStyle);

        cellName_points.setCellValue(namePoints.get(k).getName());
        cellThousand_t.setCellValue(namePoints.get(k).getTonn());
        cellThousand_pos_ed.setCellValue(namePoints.get(k).getUnits());
        cellThousand_sht.setCellValue(namePoints.get(k).getPieces());
        cellThousand_part.setCellValue(namePoints.get(k).getParties());
        cellThousand_m2.setCellValue(namePoints.get(k).getM2());
        cellThousand_m3.setCellValue(namePoints.get(k).getM3());
        cellRailway_wagons.setCellValue(namePoints.get(k).getWagons());
        cellMotor_transport.setCellValue(namePoints.get(k).getMotor_transport());
        cellContainer.setCellValue(namePoints.get(k).getContainers());
        cellBaggage.setCellValue(namePoints.get(k).getBaggage());
        cellAirplane.setCellValue(namePoints.get(k).getAirplane());

    }
}
