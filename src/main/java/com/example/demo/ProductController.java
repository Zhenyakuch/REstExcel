package com.example.demo;

import com.aspose.cells.*;
import com.example.demo.model.TotalMass;
import com.example.demo.model.dto.CountryReport;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellRange;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;

@RestController
@Slf4j
public class ProductController {

    @PostMapping("/product")
    public String printData(@RequestBody CountryReport countryRow) throws Exception {
        log.debug("CountryReport "+ countryRow);

        System.out.println(countryRow.getProduct());
        System.out.println(countryRow.getStarDate());
        System.out.println(countryRow.getEndDate());
        System.out.println(countryRow.getCountryRows());

        System.out.println();
        TotalMass totalMass= new TotalMass(countryRow);
        log.debug("TotalMass "+ totalMass);


        Workbook workbook = new Workbook("src/main/resources/Ввоз.xlsx");
        Worksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.replace("product", countryRow.getProduct());
        worksheet.replace("startDate", countryRow.getStarDate().toString());
        worksheet.replace("endDate", countryRow.getEndDate().toString());

        // Получить ссылку на ячейку «A1» из ячеек рабочего листа
        //Cell cell = workbook.getWorksheets().get(0).getCells().get("A4");

// Установите «Привет, мир!» значение в ячейку "A1"
       // cell.setValue(countryRow.getProduct());
       // workbook.getWorksheets().removeAt("Evaluation Warning");

//        ReplaceOptions replaceOptions = new ReplaceOptions();
//        replaceOptions.setCaseSensitive(false);
//        replaceOptions.setMatchEntireCellContents(false);
//
//        workbook.replace("product",countryRow.getProduct(),replaceOptions);
        workbook.save("src/main/resources/updated.xlsx");


return "";

        }
}