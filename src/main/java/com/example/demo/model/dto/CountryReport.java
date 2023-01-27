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
    private boolean isProduct;
    private boolean isImport;
    private boolean isReexport;
    private boolean isTranzitEAEU;
    private boolean isFlowers;
    private List<CountryRow> countryRows;
    private List<UnitsFlowers> units;
    private List<PieceFlowers> piece;
    private Fss fss;
    private int allFss2022;
}
