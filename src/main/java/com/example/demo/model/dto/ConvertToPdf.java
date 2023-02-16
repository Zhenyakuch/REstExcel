package com.example.demo.model.dto;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;

import java.io.File;

public class ConvertToPdf {

    public static String convert(String nameFile) {

        Document doc = new Document(nameFile);
        doc.saveToFile(nameFile + ".pdf", FileFormat.PDF);

        return ConBase64.convert(new File(nameFile + ".pdf"));
    }
}
