package com.example.demo.model.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Setter
@Getter
@ToString(callSuper = true)
public class StickerProduct {
    private String name;
    private String weight;
    private String quantity;
    private String unit;
    private String additional_info;
    private String seal_number;
}
