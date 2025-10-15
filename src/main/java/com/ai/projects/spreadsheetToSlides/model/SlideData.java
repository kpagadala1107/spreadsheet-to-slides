package com.ai.projects.spreadsheetToSlides.model;


import java.util.List;


public class SlideData {
    private String title;
    private List<String> content;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public List<String> getContent() {
        return content;
    }

    public void setContent(List<String> content) {
        this.content = content;
    }

    @Override
    public String toString() {
        return "SlideData{" +
                "title='" + title + '\'' +
                ", content=" + content +
                '}';
    }
}