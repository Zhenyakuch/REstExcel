package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Setter
@Getter
@ToString(callSuper = true)

public class NamePoints {
    private String name;
    private int tonn; //тонны
    private int units;//пос.ед.
    private int pieces;//штуки
    private int parties;//партии
    private int m2_packages;
    private int m3;
    private int wagons;
    private int motor_transport;
    private int containers;
    private int baggage;
    private int airplane;

}
