package com.ai.projects.spreadsheetToSlides.service;

import com.ai.projects.spreadsheetToSlides.model.SlideData;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.reactive.function.client.WebClient;
import reactor.core.publisher.Mono;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class ConversionServiceUpdated {

    @Value("${openai.api.key}")
    private String openaiApiKey;

    private final WebClient webClient = WebClient.create("https://api.openai.com");

    public byte[] convertToPpt(MultipartFile file, String prompt) throws Exception {
        // Parse spreadsheet data for charts
        Map<String, Object> spreadsheetData = parseSpreadsheetData(file);

        // LLM prompt with chart suggestions
        String promptToLlm = prompt + spreadsheetData.get("textData");

        // Call OpenAI
        String llmResponse = callOpenAI(promptToLlm);
        List<SlideData> slides = parseSlides(llmResponse);

        // Generate PPT with charts
        XMLSlideShow ppt = new XMLSlideShow();

        for (SlideData slideData : slides) {
            XSLFSlide slide = ppt.createSlide();

            // Create title
            createTitle(slide, slideData.getTitle());

            // Create content and charts
            if (slideData.getChartType() != null) {
                createChart(slide, slideData, spreadsheetData);
            } else {
                createTextContent(slide, slideData);
            }
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);
        ppt.close();
        return out.toByteArray();
    }

    private Map<String, Object> parseSpreadsheetData(MultipartFile file) throws Exception {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());

        Map<String, Object> result = new HashMap<>();
        List<Map<String, Object>> allSheetsData = new ArrayList<>();
        List<String> textData = new ArrayList<>();
        List<String> categories = new ArrayList<>();
        List<Double> values = new ArrayList<>();
        Map<String, List<Double>> seriesData = new HashMap<>();

        // Process all sheets
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            Map<String, Object> sheetData = processSheet(sheet);
            allSheetsData.add(sheetData);

            // Aggregate data from all sheets
            @SuppressWarnings("unchecked")
            List<String> sheetText = (List<String>) sheetData.get("textData");
            textData.addAll(sheetText);

            @SuppressWarnings("unchecked")
            List<String> sheetCategories = (List<String>) sheetData.get("categories");
            @SuppressWarnings("unchecked")
            List<Double> sheetValues = (List<Double>) sheetData.get("values");

            if (!sheetCategories.isEmpty() && !sheetValues.isEmpty()) {
                categories.addAll(sheetCategories);
                values.addAll(sheetValues);
            }

            @SuppressWarnings("unchecked")
            Map<String, List<Double>> sheetSeries = (Map<String, List<Double>>) sheetData.get("seriesData");
            seriesData.putAll(sheetSeries);
        }

        workbook.close();

        result.put("textData", String.join("\n", textData));
        result.put("categories", categories);
        result.put("values", values);
        result.put("seriesData", seriesData);
        result.put("allSheetsData", allSheetsData);
        result.put("chartRecommendations", generateChartRecommendations(categories, values, seriesData));

        return result;
    }

    private Map<String, Object> processSheet(Sheet sheet) {
        List<String> textData = new ArrayList<>();
        List<String> categories = new ArrayList<>();
        List<Double> values = new ArrayList<>();
        Map<String, List<Double>> seriesData = new HashMap<>();
        List<String> headers = new ArrayList<>();
        Map<String, Integer> columnTypes = new HashMap<>();

        boolean isFirstRow = true;
        int maxColumns = 0;

        // First pass: determine structure and headers
        for (Row row : sheet) {
            if (row.getLastCellNum() > maxColumns) {
                maxColumns = row.getLastCellNum();
            }

            if (isFirstRow) {
                // Extract headers
                for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    String headerValue = cell != null ? cell.toString().trim() : "Column" + (cellIndex + 1);
                    headers.add(headerValue);

                    // Initialize column type detection
                    columnTypes.put(headerValue, 0); // 0=unknown, 1=text, 2=numeric, 3=mixed
                }
                isFirstRow = false;
                continue;
            }

            // Analyze column types
            for (int cellIndex = 0; cellIndex < Math.min(row.getLastCellNum(), headers.size()); cellIndex++) {
                Cell cell = row.getCell(cellIndex);
                if (cell != null && headers.size() > cellIndex) {
                    String header = headers.get(cellIndex);
                    int currentType = columnTypes.get(header);

                    if (cell.getCellType() == CellType.NUMERIC) {
                        columnTypes.put(header, currentType == 1 ? 3 : 2); // numeric or mixed
                    } else if (cell.getCellType() == CellType.STRING && !cell.getStringCellValue().trim().isEmpty()) {
                        columnTypes.put(header, currentType == 2 ? 3 : 1); // text or mixed
                    }
                }
            }
        }

        // Second pass: extract data based on detected structure
        isFirstRow = true;
        for (Row row : sheet) {
            if (isFirstRow) {
                isFirstRow = false;
                continue;
            }

            StringBuilder rowText = new StringBuilder();
            Map<String, Object> rowData = new HashMap<>();
            String primaryCategory = null;
            Double primaryValue = null;

            for (int cellIndex = 0; cellIndex < Math.min(row.getLastCellNum(), headers.size()); cellIndex++) {
                Cell cell = row.getCell(cellIndex);
                String header = headers.get(cellIndex);
                String cellValue = "";

                if (cell != null) {
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                cellValue = cell.getDateCellValue().toString();
                            } else {
                                double numValue = cell.getNumericCellValue();
                                cellValue = String.valueOf(numValue);
                                rowData.put(header, numValue);

                                // Store for series data
                                seriesData.computeIfAbsent(header, k -> new ArrayList<>()).add(numValue);

                                // Set primary value if this is the first numeric column
                                if (primaryValue == null && columnTypes.get(header) == 2) {
                                    primaryValue = numValue;
                                }
                            }
                            break;
                        case STRING:
                            cellValue = cell.getStringCellValue().trim();
                            rowData.put(header, cellValue);

                            // Set primary category if this is the first text column
                            if (primaryCategory == null && columnTypes.get(header) == 1 && !cellValue.isEmpty()) {
                                primaryCategory = cellValue;
                            }
                            break;
                        case BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            rowData.put(header, cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            try {
                                cellValue = String.valueOf(cell.getNumericCellValue());
                                rowData.put(header, cell.getNumericCellValue());
                            } catch (Exception e) {
                                cellValue = cell.getCellFormula();
                                rowData.put(header, cellValue);
                            }
                            break;
                        default:
                            cellValue = cell.toString();
                            break;
                    }
                }

                rowText.append(cellValue).append(" ");
            }

            String rowTextStr = rowText.toString().trim();
            if (!rowTextStr.isEmpty()) {
                textData.add(rowTextStr);
            }

            // Add to categories and values for basic charting
            if (primaryCategory != null && primaryValue != null) {
                categories.add(primaryCategory);
                values.add(primaryValue);
            }
        }

        Map<String, Object> sheetResult = new HashMap<>();
        sheetResult.put("textData", textData);
        sheetResult.put("categories", categories);
        sheetResult.put("values", values);
        sheetResult.put("seriesData", seriesData);
        sheetResult.put("headers", headers);
        sheetResult.put("columnTypes", columnTypes);
        sheetResult.put("sheetName", sheet.getSheetName());

        return sheetResult;
    }

    private List<String> generateChartRecommendations(List<String> categories, List<Double> values, Map<String, List<Double>> seriesData) {
        List<String> recommendations = new ArrayList<>();

        // Basic recommendations based on data structure
        if (!categories.isEmpty() && !values.isEmpty()) {
            if (categories.size() <= 6) {
                recommendations.add("pie chart for " + categories.size() + " categories");
            }
            recommendations.add("bar chart for categorical comparison");

            if (categories.size() > 10) {
                recommendations.add("line chart for trend analysis");
            }
        }

        // Advanced recommendations based on series data
        if (seriesData.size() > 1) {
            recommendations.add("multi-series bar chart for comparison");
            recommendations.add("stacked bar chart for composition");
        }

        // Time series detection
        boolean hasTimeData = seriesData.keySet().stream()
                .anyMatch(header -> header.toLowerCase().contains("date") ||
                        header.toLowerCase().contains("time") ||
                        header.toLowerCase().contains("year") ||
                        header.toLowerCase().contains("month"));

        if (hasTimeData) {
            recommendations.add("line chart for time series data");
            recommendations.add("area chart for cumulative trends");
        }

        return recommendations;
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

    private void createChart(XSLFSlide slide, SlideData slideData, Map<String, Object> data) {
        @SuppressWarnings("unchecked")
        List<String> categories = (List<String>) data.get("categories");
        @SuppressWarnings("unchecked")
        List<Double> values = (List<Double>) data.get("values");

        if (categories.isEmpty() || values.isEmpty()) {
            createTextContent(slide, slideData);
            return;
        }

        // Create a text-based chart representation for now
        createChartAsText(slide, slideData.getChartType(), categories, values);
    }

    private void createChartAsText(XSLFSlide slide, String chartType, List<String> categories, List<Double> values) {
        XSLFTextBox chartShape = slide.createTextBox();
        chartShape.setAnchor(new Rectangle(50, 150, 600, 300));

        XSLFTextParagraph titlePara = chartShape.addNewTextParagraph();
        XSLFTextRun titleRun = titlePara.addNewTextRun();
        titleRun.setText(chartType.toUpperCase() + " CHART DATA:");
        titleRun.setFontSize(18.0);
        titleRun.setBold(true);

        for (int i = 0; i < Math.min(categories.size(), values.size()); i++) {
            XSLFTextParagraph dataPara = chartShape.addNewTextParagraph();
            XSLFTextRun dataRun = dataPara.addNewTextRun();
            dataRun.setText(String.format("• %s: %.2f", categories.get(i), values.get(i)));
            dataRun.setFontSize(14.0);
        }
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
