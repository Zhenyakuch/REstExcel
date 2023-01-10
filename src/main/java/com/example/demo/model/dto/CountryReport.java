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

    private int fss2021Brest;
    private int fss2021Vitebsk;
    private int fss2021Gomel;
    private int fss2021Grodno;
    private int fss2021Minsk;
    private int fss2021Mogilev;
    private int fss2022Brest;
    private int fss2022Vitebsk;
    private int fss2022Gomel;
    private int fss2022Grodno;
    private int fss2022Minsk;
    private int fss2022Mogilev;
    private int fss2021Brest_7;
    private int fss2021Vitebsk_7;
    private int fss2021Gomel_7;
    private int fss2021Grodno_7;
    private int fss2021Minsk_7;
    private int fss2021Mogilev_7;
    private int fss2022Brest_7;
    private int fss2022Vitebsk_7;
    private int fss2022Gomel_7;
    private int fss2022Grodno_7;
    private int fss2022Minsk_7;
    private int fss2022Mogilev_7;

    private int fssPeriod2022_21;
    private int fssPeriod2021_21;
    private int fssPeriod7_21;

    private int fssPeriod2022_22;
    private int fssPeriod2021_22;
    private int fssPeriod7_22;
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

    private int unit_material;
    private int unit_material_7;
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
