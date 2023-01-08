package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;
import java.util.List;

@Getter
@Setter
@ToString
public class CountryReport {

    //переменные json импорта экспорта и реэк
    private LocalDate starDate;
    private LocalDate endDate;
    private String reqCountryOrProduct;
    private String resCountryOrProduct;
    private boolean isImport;
    private List<CountryRow> countryRows;
}
