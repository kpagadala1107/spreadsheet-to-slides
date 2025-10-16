package com.ai.projects.spreadsheetToSlides.service;

import com.ai.projects.spreadsheetToSlides.model.SlideData;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.reactive.function.client.WebClient;
import reactor.core.publisher.Mono;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class ConversionService {

    @Value("${openai.api.key}")
    private String openaiApiKey;

    private final WebClient webClient = WebClient.create("https://api.openai.com");

    public byte[] convertToPpt(MultipartFile file, String targetAudience) throws Exception {
        Map<String, Object> metadata = parseSpreadsheetMetadata(file);

//        targetAudience = "Project Managers";

        // Build prompt for LLM
        StringBuilder promptBuilder = new StringBuilder();
        promptBuilder.append("Create presentation slides for the following Excel file, targeting audience: ")
                .append(targetAudience).append(".\n");
        List<Map<String, Object>> sheetsInfo = (List<Map<String, Object>>) metadata.get("sheetsInfo");
        for (Map<String, Object> sheet : sheetsInfo) {
            promptBuilder.append("Sheet: ").append(sheet.get("sheetName")).append("\nHeaders: ");
            promptBuilder.append(String.join(", ", (List<String>) sheet.get("headers"))).append("\n");
        }
        String promptToLlm = promptBuilder.toString();

        // Call LLM and generate slides as before
        String llmResponse = callOpenAI(promptToLlm);
        List<SlideData> slides = parseSlides(llmResponse);

        XMLSlideShow ppt = new XMLSlideShow();
        for (SlideData slideData : slides) {
            XSLFSlide slide = ppt.createSlide();
            createTitle(slide, slideData.getTitle());
            createTextContent(slide, slideData);
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);
        ppt.close();
        return out.toByteArray();
    }

    private Map<String, Object> parseSpreadsheetMetadata(MultipartFile file) throws Exception {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        List<Map<String, Object>> sheetsInfo = new ArrayList<>();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            Map<String, Object> sheetInfo = new HashMap<>();
            sheetInfo.put("sheetName", sheet.getSheetName());

            List<String> headers = new ArrayList<>();
            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                    Cell cell = headerRow.getCell(j);
                    headers.add(cell != null ? cell.toString().trim() : "Column" + (j + 1));
                }
            }
            sheetInfo.put("headers", headers);
            sheetsInfo.add(sheetInfo);
        }
        workbook.close();

        Map<String, Object> result = new HashMap<>();
        result.put("sheetsInfo", sheetsInfo);
        return result;
    }

    private List<SlideData> parseSlides(String response) {
        List<SlideData> slides = new ArrayList<>();
        String[] sections = response.split("\\n\\n");

        for (String section : sections) {
            if (section.trim().isEmpty()) continue;

            SlideData slide = new SlideData();
            String[] lines = section.split("\\n");

            if (lines.length > 0) {
                slide.setTitle(lines[0].replaceAll("^#+\\s*", "").trim());

                List<String> content = new ArrayList<>();
                String chartType = null;

                for (int i = 1; i < lines.length; i++) {
                    String line = lines[i].trim();
                    if (!line.isEmpty()) {
                        // Check for chart suggestions
                        if (line.toLowerCase().contains("pie chart")) {
                            chartType = "pie";
                        } else if (line.toLowerCase().contains("bar chart")) {
                            chartType = "bar";
                        } else if (line.toLowerCase().contains("line chart")) {
                            chartType = "line";
                        } else {
                            content.add(line.replaceAll("^[•\\-\\*]\\s*", ""));
                        }
                    }
                }

                slide.setContent(content.isEmpty() ? List.of("No content available") : content);
                slide.setChartType(chartType);
            }

            slides.add(slide);
        }

        if (slides.isEmpty()) {
            SlideData fallbackSlide = new SlideData();
            fallbackSlide.setTitle("Data Summary");
            fallbackSlide.setContent(List.of("Data visualization"));
            fallbackSlide.setChartType("pie");
            slides.add(fallbackSlide);
        }

        return slides;
    }

    private void createTitle(XSLFSlide slide, String title) {
        XSLFTextBox titleShape = slide.createTextBox();
        titleShape.setAnchor(new Rectangle(50, 20, 600, 60));
        XSLFTextParagraph titleParagraph = titleShape.addNewTextParagraph();
        XSLFTextRun titleRun = titleParagraph.addNewTextRun();
        titleRun.setText(title);
        titleRun.setFontSize(24.0);
        titleRun.setBold(true);
    }

    private void createTextContent(XSLFSlide slide, SlideData slideData) {
        XSLFTextBox contentShape = slide.createTextBox();
        contentShape.setAnchor(new Rectangle(50, 100, 600, 400));

        for (String contentLine : slideData.getContent()) {
            XSLFTextParagraph paragraph = contentShape.addNewTextParagraph();
            XSLFTextRun run = paragraph.addNewTextRun();
            run.setText("• " + contentLine);
            run.setFontSize(16.0);
            paragraph.setLeftMargin(20.0);
        }
    }

    private String callOpenAI(String prompt) throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        Map<String, Object> request = Map.of(
                "model", "gpt-3.5-turbo",
//                "model", "gpt-4-turbo",
                "messages", List.of(Map.of("role", "user", "content", prompt)),
                "max_tokens", 1000
        );
        String requestBody = mapper.writeValueAsString(request);

        Mono<String> response = webClient.post()
                .uri("/v1/chat/completions")
                .header("Authorization", "Bearer " + openaiApiKey)
                .contentType(MediaType.APPLICATION_JSON)
                .bodyValue(requestBody)
                .retrieve()
                .bodyToMono(String.class);

        String result = response.block();

        // Extract content from OpenAI response
        ObjectMapper responseMapper = new ObjectMapper();
        Map<String, Object> responseMap = responseMapper.readValue(result, Map.class);
        List<Map<String, Object>> choices = (List<Map<String, Object>>) responseMap.get("choices");
        Map<String, Object> message = (Map<String, Object>) choices.get(0).get("message");

        return (String) message.get("content");
    }

}
