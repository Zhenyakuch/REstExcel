package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.List;

@Getter
@Setter
@ToString
public class CountryRow {

    //переменные json импорта экспорта и реэк
    private String resCountryOrProduct;
    private ElementMass massProduct;
    private List<ElementRegion> regions;

}
