package com.ai.projects.spreadsheetToSlides.configuration;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.cors.CorsConfiguration;
import org.springframework.web.cors.UrlBasedCorsConfigurationSource;
import org.springframework.web.filter.CorsFilter;

import java.util.List;

@Configuration
public class CorsConfig {

    @Bean
    public CorsFilter corsFilter() {
        UrlBasedCorsConfigurationSource source = new UrlBasedCorsConfigurationSource();
        CorsConfiguration config = new CorsConfiguration();

        // Allow your React app's origin - removed trailing slashes and added more flexibility
        config.setAllowedOrigins(List.of(
            "http://localhost:3000", "https://spreadsheet-to-slides-ui.netlify.app"
        ));

        // Allow all needed methods including OPTIONS for preflight
        config.setAllowedMethods(List.of("GET", "POST", "PUT", "DELETE", "OPTIONS", "PATCH"));

        // Allow all headers
        config.setAllowedHeaders(List.of("*"));

        // Allow credentials if needed
        config.setAllowCredentials(true);

        // Expose common headers
        config.setExposedHeaders(List.of("Authorization", "Content-Type"));

        // Set max age for preflight cache
        config.setMaxAge(3600L);

        // Apply this configuration to all paths
        source.registerCorsConfiguration("/**", config);

        return new CorsFilter(source);
    }
}