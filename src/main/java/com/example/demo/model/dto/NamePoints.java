package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Setter
@Getter
@ToString(callSuper = true)

public class NamePoints {
    private String name;
    private int tonn;
    private int units;
    private int pieces;
    private int parties;
    private int m2;
    private int m3;
    private int wagons;
    private int motor_transport;
    private int containers;
    private int baggage;
    private int airplane;

}
