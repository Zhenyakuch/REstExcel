package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Setter
@Getter
@ToString(callSuper = true)
public class ElementRegion extends ElementMass {

    //переменные json импорта экспорта и реэк
    private int region;


    //переменные json транзит

    //private List<ElementRegion> reg;
    private String name_points;
    private int thousand_t;
    private int thousand_pos_ed;
    private int thousand_sht;
    private int thousand_part;
    private int thousand_m2;
    private int thousand_m3;
    private int railway_wagons;
    private int motor_transport;
    private int container;
    private int baggage;
    private int airplane;


}
