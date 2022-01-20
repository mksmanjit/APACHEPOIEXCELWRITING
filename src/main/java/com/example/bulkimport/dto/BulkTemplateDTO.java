package com.example.bulkimport.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Data;

import java.io.Serializable;
import java.util.List;

@Data
@JsonIgnoreProperties(ignoreUnknown = true)
public class BulkTemplateDTO implements Serializable {
    private String categoryName;
    private Long categoryId;
    private List<ColumnDTO> fixedColumn;
    private List<ColumnDTO> customFields;
    private List<ColumnDTO> productOptions;

    @Data
    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class ColumnDTO implements Serializable {
        private String columnName;
        private String type;
        private String description;
        private List<String> allowedValues;
    }

}
