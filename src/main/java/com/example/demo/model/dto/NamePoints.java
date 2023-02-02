package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Setter
@Getter
@ToString(callSuper = true)

public class NamePoints {
    private String name;
    private double tonn; //тонны
    private double units;//пос.ед.
    private double pieces;//штуки
    private double parties;//партии
    private double m2_packages;
    private double m3;
    private double wagons;
    private double motor_transport;
    private double containers;
    private double baggage;
    private double airplane;

}
