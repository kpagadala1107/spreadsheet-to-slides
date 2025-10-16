package com.ai.projects.spreadsheetToSlides.rest;

import com.ai.projects.spreadsheetToSlides.service.ConversionService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api")
public class ConversionController {

    private final ConversionService conversionService;

    public ConversionController(ConversionService conversionService) {
        this.conversionService = conversionService;
    }

    @PostMapping("/convert")
    public ResponseEntity<byte[]> convertSpreadsheet(@RequestParam("file") MultipartFile file, @RequestParam("targetAudience") String targetAudience) throws Exception {
        byte[] pptData = conversionService.convertToPpt(file, targetAudience);
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "presentation.pptx");
        return ResponseEntity.ok().headers(headers).body(pptData);
    }
}