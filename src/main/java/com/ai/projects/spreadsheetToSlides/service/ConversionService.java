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
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Service
public class ConversionService {

    @Value("${openai.api.key}")
    private String openaiApiKey;

    private final WebClient webClient = WebClient.create("https://api.openai.com");

    public byte[] convertToPpt(MultipartFile file) throws Exception {
        // Parse spreadsheet
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);
        List<String> dataSummary = new ArrayList<>();
        for (Row row : sheet) {
            StringBuilder rowData = new StringBuilder();
            for (Cell cell : row) {
                rowData.append(cell.toString()).append(" ");
            }
            dataSummary.add(rowData.toString().trim());
        }
        workbook.close();

        // LLM prompt
//        String prompt = "Summarize this spreadsheet data into a PowerPoint slide outline with titles, bullets, and visualization suggestions: " + String.join("\n", dataSummary);
        String prompt = "I need to create a professional PowerPoint presentation based on the data in this spreadsheet. The presentation should focus on summarizing key metrics, highlighting trends, and emphasizing important insights. Please generate slide titles, bullet points, and data-driven narratives that can clearly convey the analysis to a business audience. Include suggestions for relevant charts or visuals to accompany the data: " + String.join("\n", dataSummary);

        // Call OpenAI
        String llmResponse = callOpenAI(prompt);
        List<SlideData> slides = parseSlides(llmResponse);

        // Generate PPT with proper positioning
        XMLSlideShow ppt = new XMLSlideShow();
        for (SlideData slideData : slides) {
            XSLFSlide slide = ppt.createSlide();

            // Create and position title
            XSLFTextBox titleShape = slide.createTextBox();
            titleShape.setAnchor(new Rectangle(50, 50, 600, 80));
            XSLFTextParagraph titleParagraph = titleShape.addNewTextParagraph();
            XSLFTextRun titleRun = titleParagraph.addNewTextRun();
            titleRun.setText(slideData.getTitle());
            titleRun.setFontSize(24.0);
            titleRun.setBold(true);

            // Create and position content
            XSLFTextBox contentShape = slide.createTextBox();
            contentShape.setAnchor(new Rectangle(50, 150, 600, 400));

            for (String contentLine : slideData.getContent()) {
                XSLFTextParagraph paragraph = contentShape.addNewTextParagraph();
                XSLFTextRun run = paragraph.addNewTextRun();
                run.setText("• " + contentLine);
                run.setFontSize(16.0);
                paragraph.setLeftMargin(20.0);
            }
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);
        ppt.close();
        return out.toByteArray();
    }

    private String callOpenAI(String prompt) throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        Map<String, Object> request = Map.of(
//                "model", "gpt-3.5-turbo",
                "model", "gpt-4-turbo",
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

    private List<SlideData> parseSlides(String response) {
        List<SlideData> slides = new ArrayList<>();

        // Enhanced parsing - split response into logical sections
        String[] sections = response.split("\\n\\n");

        for (String section : sections) {
            if (section.trim().isEmpty()) continue;

            SlideData slide = new SlideData();
            String[] lines = section.split("\\n");

            if (lines.length > 0) {
                slide.setTitle(lines[0].replaceAll("^#+\\s*", "").trim());

                List<String> content = new ArrayList<>();
                for (int i = 1; i < lines.length; i++) {
                    String line = lines[i].trim();
                    if (!line.isEmpty()) {
                        content.add(line.replaceAll("^[•\\-\\*]\\s*", ""));
                    }
                }
                slide.setContent(content.isEmpty() ? List.of("No content available") : content);
            }

            slides.add(slide);
        }

        // Fallback if no slides were parsed
        if (slides.isEmpty()) {
            SlideData fallbackSlide = new SlideData();
            fallbackSlide.setTitle("Data Summary");
            fallbackSlide.setContent(List.of(response.length() > 100 ? response.substring(0, 100) + "..." : response));
            slides.add(fallbackSlide);
        }

        return slides;
    }
}
