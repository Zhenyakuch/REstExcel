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

    private int allFss2022;


    private int allFss;
    private int allFss_7;
    private int allFssBrest;
    private int allFss_7Brest;
    private int allFssVitebsk;
    private int allFss_7Vitebsk;
    private int allFssGomel;
    private int allFss_7Gomel;
    private int allFssGrodno;
    private int allFss_7Grodno;
    private int allFssMinsk;
    private int allFss_7Minsk;
    private int allFssMogilev;
    private int allFss_7Mogilev;

    private int unit_materialFromTo;
    private int unit_material_7;
    private int unit_material_2022;
    private int unit_materialBrest;
    private int unit_material_7Brest;
    private int unit_materialVitebsk;
    private int unit_material_7Vitebsk;
    private int unit_materialGomel;
    private int unit_material_7Gomel;
    private int unit_materialGrodno;
    private int unit_material_7Grodno;
    private int unit_materialMinsk;
    private int unit_materials_7Minsk;
    private int unit_materialMogilev;
    private int unit_material_7Mogilev;
    private int piece_material;
    private int piece_material_7;
    private int piece_materialBrest;
    private int piece_material_7Brest;
    private int piece_materialVitebsk;
    private int piece_material_7Vitebsk;
    private int piece_materialGomel;
    private int piece_material_7Gomel;
    private int piece_materialGrodno;
    private int piece_material_7Grodno;
    private int piece_materialMinsk;
    private int piece_materials_7Minsk;
    private int piece_materialMogilev;
    private int piece_material_7Mogilev;


}
