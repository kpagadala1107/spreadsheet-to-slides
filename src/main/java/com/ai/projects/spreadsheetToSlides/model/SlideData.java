package com.ai.projects.spreadsheetToSlides.model;

import java.util.List;

public class SlideData {
    private String title;
    private List<String> content;
    private String chartType; // "pie", "bar", "line", or null

    // Constructors, getters, and setters
    public SlideData() {}

    public String getTitle() { return title; }
    public void setTitle(String title) { this.title = title; }

    public List<String> getContent() { return content; }
    public void setContent(List<String> content) { this.content = content; }

    public String getChartType() { return chartType; }
    public void setChartType(String chartType) { this.chartType = chartType; }
}
