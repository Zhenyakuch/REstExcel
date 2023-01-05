package com.example.demo.model.dto;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Setter
@Getter
@ToString(callSuper = true)
public class ElementRegion extends ElementMass{
    private int region;
    private int date2021;
    private int date2022;
}
