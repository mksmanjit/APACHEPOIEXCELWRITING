package com.example.bulkimport.controller;

import com.example.bulkimport.service.BulkTemplateDataService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.util.HashSet;
import java.util.List;

@RestController
@RequestMapping(value = "/bulk-export", produces = {MediaType.APPLICATION_JSON_VALUE, MediaType.APPLICATION_XML_VALUE})
public class BulkExportController {

    @Autowired
    private BulkTemplateDataService bulkTemplateDataService;

    @GetMapping(value = "/download-template",produces = "application/octet-stream")
    public ResponseEntity<Resource> getStockReport(@RequestParam("catIds") List<Long> categoryIds) {
        return generateResponse(bulkTemplateDataService.generateReport(new HashSet<>(categoryIds)));
    }

    ResponseEntity<Resource> generateResponse(ByteArrayResource resource) {
        String reportName = "Bulk_Import_multiple_category_template.xlsx";
        HttpHeaders headers = new HttpHeaders();
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + reportName);
        return ResponseEntity.ok()
                .headers(headers)
                .contentLength(resource.contentLength())
                .contentType(MediaType.parseMediaType("application/octet-stream"))
                .body(resource);
    }
}
