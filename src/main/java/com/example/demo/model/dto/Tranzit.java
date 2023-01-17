package com.example.demo.model.dto;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class Tranzit {

    public static void createRows(XSSFWorkbook xssfWorkbook, XSSFCellStyle cellStyle, List<ElementRegion> regions, List<NamePoints> namePoints) {

        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);

        int rowLast = sheet.getLastRowNum();
        XSSFRow row = sheet.createRow(rowLast + 1);

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
        sheet.addMergedRegion(new CellRangeAddress(rowLast + 1, rowLast + 1, 0, 11));
//
//
//
        XSSFRow row2 = sheet.createRow(rowLast + 2);
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

        for (int k = 0; regions.size() > k; k++) {

            switch (regions.get(k).getRegion()) {
                case 1:
                    cellName_obl.setCellValue("БРЕСТСКАЯ ОБЛАСТЬ");
                    cellName_points.setCellValue(namePoints.get(k).getName());
                    break;
            }
        }
    }

}
