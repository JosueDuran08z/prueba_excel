package com.example.prueba_excel.config;

import org.springframework.context.annotation.Configuration;

import jakarta.annotation.PostConstruct;

@Configuration
public class FontConfiguration {

    @PostConstruct
    public void setProperty() {
        System.setProperty("org.apache.poi.ss.ignoreMissingFontSystem", "true");
    }
}
