package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.List;

@Setter
@Getter
@ToString(callSuper = true)
public class ElementRegion extends ElementMass {

    //переменные json импорта экспорта и реэк
    private int region;
    private List<NamePoints> namePoints;


}
