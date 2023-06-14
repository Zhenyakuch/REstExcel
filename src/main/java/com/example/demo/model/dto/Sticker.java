package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;
import java.util.Date;
import java.util.List;

@Getter
@Setter
@ToString
public class Sticker {

    private List<StickerProduct> stickerProducts;
    private String number;
    private String origin;
    private String place;
    private String recipient;
    private String appointment;
    private double area;
    private String external_sings;
    private String provisional_definition;
    private String position;
    private String date;
    private String fio1;
    private String fio2;

}
