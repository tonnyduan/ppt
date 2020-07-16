package com.example.demo;


import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Component
@ConfigurationProperties(prefix = "ppt")
public class Config {

    private HashMap<String, String> path;
    private List<Map<String, String>> pages;

    public HashMap<String, String> getPath() {
        return path;
    }

    public void setPath(HashMap<String, String> path) {
        this.path = path;
    }

    public List<Map<String, String>> getPages() {
        return pages;
    }

    public void setPages(List<Map<String, String>> pages) {
        this.pages = pages;
    }

    @Override
    public String toString() {
        return "Config{" +
                "path=" + path +
                ", pages=" + pages +
                '}';
    }
}
