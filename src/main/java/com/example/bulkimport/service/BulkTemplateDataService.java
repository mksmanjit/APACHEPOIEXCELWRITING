package com.example.bulkimport.service;

import com.example.bulkimport.dto.BulkTemplateDTO;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.apachecommons.CommonsLog;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

@CommonsLog
@Service
public class BulkTemplateDataService {

    @Value("${stock.example.file.location}")
    private String stockExmapleFilePath;

    private Resource validJsonResource = new ClassPathResource("data/data.json");

    private static final String DOLLAR = "$";
    private static final String COLON = ":";
    private static final String EXCLAMATION = "!";

    public static final String FIELD_TYPE_FIELDS = "F";
    public static final String FIELD_TYPE_CUSTOM_FIELDS = "CF";
    public static final String FIELD_TYPE_PRODUCT_OPTIONS = "PO";
    public static final String PRODUCT_SHEET = "Fill Your Products Here";
    public static final String COLUMN_SHOW_HIDE_SHEET = "columnShowHideMap";
    public static final String DATA_SHEET = "Data";

    public ByteArrayResource generateReport(Set<Long> categoryIds) {
        ObjectMapper objectMapper = new ObjectMapper();
        try (Workbook workbook = new XSSFWorkbook(getTempFile())) {
            List<BulkTemplateDTO> bulkTemplateDTOs = objectMapper.readValue(validJsonResource.getFile(), new TypeReference<List<BulkTemplateDTO>>(){});
            List<BulkTemplateDTO> mappedCategoryBulkTemplateDTOs = bulkTemplateDTOs.stream().filter(dto -> categoryIds.contains(dto.getCategoryId())).collect(Collectors.toList());
            if (!CollectionUtils.isEmpty(mappedCategoryBulkTemplateDTOs)) {
                Sheet dataSheet = workbook.createSheet("Data");
                Map<String, List<String>> lookUpIndexInfoMap = new HashMap<>();
                populateDataSheet(dataSheet, mappedCategoryBulkTemplateDTOs, lookUpIndexInfoMap);

                Sheet columnShowHideSheet = workbook.createSheet(COLUMN_SHOW_HIDE_SHEET);
                populateColumnsShowHideMap(columnShowHideSheet, mappedCategoryBulkTemplateDTOs);

                Sheet productSheet = workbook.createSheet(PRODUCT_SHEET);
                populateProductSheet(productSheet, mappedCategoryBulkTemplateDTOs, lookUpIndexInfoMap);
                XSSFSheet sheet = ((XSSFSheet)productSheet);
                sheet.lockFormatCells(true);
                sheet.lockFormatColumns(false);
                sheet.enableLocking();

            }
            //Arranging the sheets
            workbook.setSheetOrder("Welcome", 0);
            workbook.setSheetOrder(PRODUCT_SHEET, 1);
            workbook.setSheetOrder(COLUMN_SHOW_HIDE_SHEET, 2);

            return getByteArrayResource(workbook);
        } catch (Exception e) {
            throw new RuntimeException("Error generating the product stock  template", e);
        }
    }

    private void populateColumnsShowHideMap(Sheet dataSheet, List<BulkTemplateDTO> mappedCategoryBulkTemplateDTOs) {
        Set<String> allColumns = new HashSet<>();
        Map<String, Set<String>> categoryColumnMap = new HashMap<>();

        for (BulkTemplateDTO bulkTemplateDTO : mappedCategoryBulkTemplateDTOs) {
            Set<String> columns = new HashSet<>();

            // columns.addAll(bulkTemplateDTO.getFixedColumn().stream().map(column -> column.getColumnName()).collect(Collectors.toSet()));
            columns.addAll(bulkTemplateDTO.getCustomFields().stream().map(column -> column.getColumnName()).collect(Collectors.toSet()));
            columns.addAll(bulkTemplateDTO.getProductOptions().stream().map(column -> column.getColumnName()).collect(Collectors.toSet()));
            categoryColumnMap.put(bulkTemplateDTO.getCategoryName(), columns);
            allColumns.addAll(columns);
        }

        int rowIndex = 0;
        int columnIndex = 0;



        Row dataRow = dataSheet.getRow(rowIndex) != null ? dataSheet.getRow(rowIndex) : dataSheet.createRow(rowIndex);
        Cell nextCell = dataRow.getCell(columnIndex) != null ? dataRow.getCell(columnIndex) : dataRow.createCell(columnIndex);
        nextCell.setCellValue("Columns");
        columnIndex++;
        Cell firstCell = null;
        for (String cat : categoryColumnMap.keySet()) {
            nextCell = dataRow.getCell(columnIndex) != null ? dataRow.getCell(columnIndex) : dataRow.createCell(columnIndex);
            nextCell.setCellValue(cat);
            if(firstCell == null) {
                firstCell = nextCell;
            }
            columnIndex++;
        }

        Name namedCell = dataSheet.getWorkbook().createName();
        namedCell.setNameName("categoryNames");
        namedCell.setRefersToFormula(COLUMN_SHOW_HIDE_SHEET + "!" + DOLLAR + firstCell.getAddress().toString().split("\\d+")[0] +  DOLLAR + firstCell.getAddress().toString().split("[a-zA-Z]+")[1] + ":" + DOLLAR + nextCell.getAddress().toString().split("\\d+")[0] +  DOLLAR + nextCell.getAddress().toString().split("[a-zA-Z]+")[1]);
        rowIndex++;

        firstCell = null;
        for (String column : allColumns) {
            columnIndex = 0;

            dataRow = dataSheet.getRow(rowIndex) != null ? dataSheet.getRow(rowIndex) : dataSheet.createRow(rowIndex);
            nextCell = dataRow.getCell(columnIndex) != null ? dataRow.getCell(columnIndex) : dataRow.createCell(columnIndex);
            nextCell.setCellValue(column);
            columnIndex++;

            if(firstCell == null) {
                firstCell = nextCell;
            }

            for (String cat : categoryColumnMap.keySet()) {
                nextCell = dataRow.getCell(columnIndex) != null ? dataRow.getCell(columnIndex) : dataRow.createCell(columnIndex);
                nextCell.setCellValue(categoryColumnMap.get(cat).contains(column) ? 1 : 0);
                columnIndex++;
            }

            rowIndex++;
        }

        namedCell = dataSheet.getWorkbook().createName();
        namedCell.setNameName("showHideCategoryColumn");
        namedCell.setRefersToFormula(COLUMN_SHOW_HIDE_SHEET + "!" + DOLLAR + firstCell.getAddress().toString().split("\\d+")[0] +  DOLLAR + firstCell.getAddress().toString().split("[a-zA-Z]+")[1] + ":" + DOLLAR + nextCell.getAddress().toString().split("\\d+")[0] +  DOLLAR + nextCell.getAddress().toString().split("[a-zA-Z]+")[1]);
        rowIndex++;

    }

    private File getTempFile() {
        InputStream inputStream = null;
        File tempFile = null;
        try {
            tempFile = File.createTempFile("product-import-", ".xlsx");
            File baseFile = new File(stockExmapleFilePath);
            inputStream = new FileInputStream(baseFile);
            byte[] bytes = IOUtils.toByteArray(inputStream);

            FileUtils.writeByteArrayToFile(tempFile, bytes);
        } catch (Exception e) {
            log.error("Error in copying data from base file ...", e);
        } finally {
            IOUtils.closeQuietly(inputStream);
        }
        return tempFile;
    }

    private void populateProductSheet(Sheet productSheet, List<BulkTemplateDTO> bulkTemplateDTOs, Map<String, List<String>> lookUpIndexInfoMap) {
        int columnIndex = 0;
        Set<String> columnNames = new HashSet<>();
        columnIndex = populateColumnInProductSheet(productSheet, lookUpIndexInfoMap, columnIndex, "category", columnNames,"Select the Category your product belongs to");
        for (BulkTemplateDTO bulkTemplateDTO : bulkTemplateDTOs) {
            columnIndex = populateProductSheetHeaders(productSheet, bulkTemplateDTO.getFixedColumn(), lookUpIndexInfoMap, FIELD_TYPE_FIELDS, columnIndex, columnNames);
            columnIndex = populateProductSheetHeaders(productSheet, bulkTemplateDTO.getCustomFields(), lookUpIndexInfoMap, FIELD_TYPE_CUSTOM_FIELDS, columnIndex, columnNames);
            columnIndex = populateProductSheetHeaders(productSheet, bulkTemplateDTO.getProductOptions(), lookUpIndexInfoMap, FIELD_TYPE_PRODUCT_OPTIONS, columnIndex, columnNames);
        }
    }

    private int populateProductSheetHeaders(Sheet productSheet, List<BulkTemplateDTO.ColumnDTO> columns, Map<String, List<String>> lookUpIndexInfoMap, String fieldTypeFields, int columnIndex, Set<String> columnNames) {
        for (BulkTemplateDTO.ColumnDTO column : columns) {
            if (!columnNames.contains(column.getColumnName())) {
                columnNames.add(column.getColumnName());
                columnIndex = populateColumnInProductSheet(productSheet, lookUpIndexInfoMap, columnIndex, column.getColumnName(), columnNames, column.getDescription());
            }
        }
        return columnIndex;
    }

    private int populateColumnInProductSheet(Sheet productSheet, Map<String, List<String>> lookUpIndexInfoMap, int columnIndex, String columnName, Set<String> columnNames, String description) {
        Row headerRow = productSheet.getRow(0) != null ? productSheet.getRow(0) : productSheet.createRow(0);
        XSSFCell cell = (XSSFCell) headerRow.createCell(columnIndex);
        cell.setCellValue(columnName);
        CellStyle cellStyleUnlocked = productSheet.getWorkbook().createCellStyle();
        cellStyleUnlocked.setLocked(false);
        productSheet.setDefaultColumnStyle(columnIndex, cellStyleUnlocked);
        CellStyle cellStyleLocked = productSheet.getWorkbook().createCellStyle();
        cellStyleLocked.setLocked(true);
        cell.setCellStyle(cellStyleLocked);
        CellRangeAddressList addressList = new CellRangeAddressList(0, 0, columnIndex, columnIndex);
        DataValidationHelper validationHelper = new XSSFDataValidationHelper((XSSFSheet) productSheet);
        DataValidation dataValidation = validationHelper.createValidation(new XSSFDataValidationConstraint(0,"") ,addressList);
        dataValidation.createPromptBox("", description);
        dataValidation.setShowPromptBox(true);
        productSheet.addValidationData(dataValidation);
        if(lookUpIndexInfoMap.get(columnName) != null) {

            DataValidationConstraint constraint = null;
            addressList = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL2007.getLastRowIndex(), columnIndex, columnIndex);
            String formula;
            if(columnName.equalsIgnoreCase("category")) {
                formula = "=INDIRECT(\"category\")";
            } else {
                formula = "=INDIRECT(SUBSTITUTE($A2,\" \",\"_\") &\"" + columnName.replaceAll(" ","_") + "\")";
            }

            constraint = validationHelper.createFormulaListConstraint(formula);
            dataValidation = validationHelper.createValidation(constraint, addressList);
            dataValidation.setShowErrorBox(true);
            productSheet.addValidationData(dataValidation);

        }

        if(!columnName.equalsIgnoreCase("category")) {
            SheetConditionalFormatting sheetCF = productSheet.getSheetConditionalFormatting();

            ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule("=IF(VLOOKUP(" + DOLLAR + cell.getAddress().toString().split("\\d+")[0] + DOLLAR + cell.getAddress().toString().split("[a-zA-Z]+")[1] + ",showHideCategoryColumn,MATCH($A2,categoryNames,0)+1,FALSE)>0,1,0)");
            setBackground(rule, IndexedColors.GREEN);
            //fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            addressList = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL2007.getLastRowIndex(), columnIndex, columnIndex);
            ConditionalFormattingRule[] cfRules = new ConditionalFormattingRule[]{rule};
            sheetCF.addConditionalFormatting(addressList.getCellRangeAddresses(), cfRules);


            rule = sheetCF.createConditionalFormattingRule("=IF(VLOOKUP(" + DOLLAR + cell.getAddress().toString().split("\\d+")[0] + DOLLAR + cell.getAddress().toString().split("[a-zA-Z]+")[1] + ",showHideCategoryColumn,MATCH($A2,categoryNames,0)+1,FALSE)>0,0,1)");
            PatternFormatting fill = rule.createPatternFormatting();
            fill.setFillBackgroundColor(IndexedColors.GREY_80_PERCENT.index);
            //fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            cfRules = new ConditionalFormattingRule[]{rule};
            sheetCF.addConditionalFormatting(addressList.getCellRangeAddresses(), cfRules);
        }

        List<String> mandatoryColumns = Arrays.asList("Product Name","Product Description","Image Url1");
        if(mandatoryColumns.contains(columnName)) {
            SheetConditionalFormatting sheetCF = productSheet.getSheetConditionalFormatting();

            ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule("=IF($A2 <> \"\",1,0)");
            setBackground(rule, IndexedColors.RED);
            //fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

            ConditionalFormattingRule[] cfRules = new ConditionalFormattingRule[]{rule};
            addressList = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL2007.getLastRowIndex(), columnIndex, columnIndex);

            sheetCF.addConditionalFormatting(addressList.getCellRangeAddresses(), cfRules);
        }

        columnIndex++;
        return columnIndex;
    }

    private void setBackground(ConditionalFormattingRule rule, IndexedColors green) {
        BorderFormatting fill = rule.createBorderFormatting();
        fill.setBottomBorderColor(green.index);
        fill.setBorderBottom(BorderStyle.MEDIUM);
        fill.setTopBorderColor(green.index);
        fill.setBorderTop(BorderStyle.MEDIUM);
        fill.setLeftBorderColor(green.index);
        fill.setBorderLeft(BorderStyle.MEDIUM);
        fill.setRightBorderColor(green.index);
        fill.setBorderRight(BorderStyle.MEDIUM);
    }

    public void populateDataSheet(Sheet dataSheet,List<BulkTemplateDTO> bulkTemplateDTOs,Map<String,List<String>> lookUpIndexInfoMap) {
        int columnIndex = 0;
        columnIndex = populateDataSheetForCategory(dataSheet,bulkTemplateDTOs,lookUpIndexInfoMap,FIELD_TYPE_FIELDS,columnIndex);
        for(BulkTemplateDTO bulkTemplateDTO : bulkTemplateDTOs) {

            columnIndex = populateDataSheetForFields(bulkTemplateDTO.getCategoryName() ,dataSheet, bulkTemplateDTO.getFixedColumn(), lookUpIndexInfoMap, FIELD_TYPE_FIELDS, columnIndex);
            columnIndex = populateDataSheetForFields(bulkTemplateDTO.getCategoryName() ,dataSheet, bulkTemplateDTO.getCustomFields(), lookUpIndexInfoMap, FIELD_TYPE_CUSTOM_FIELDS, columnIndex);
            columnIndex = populateDataSheetForFields(bulkTemplateDTO.getCategoryName() ,dataSheet,bulkTemplateDTO.getProductOptions(), lookUpIndexInfoMap, FIELD_TYPE_PRODUCT_OPTIONS,columnIndex);
        }

    }

    private int populateDataSheetForCategory(Sheet dataSheet, List<BulkTemplateDTO> bulkTemplateDTOs, Map<String, List<String>> lookUpIndexInfoMap, String fieldType, int columnIndex) {
        int rowIndex = 0;
        Row headerRow = dataSheet.getRow(rowIndex) != null ? dataSheet.getRow(rowIndex) : dataSheet.createRow(rowIndex);
        XSSFCell cell = (XSSFCell) headerRow.createCell(columnIndex);
        cell.setCellValue(fieldType);
        rowIndex++;

        headerRow = dataSheet.getRow(rowIndex) != null ? dataSheet.getRow(rowIndex) : dataSheet.createRow(rowIndex);
        cell = (XSSFCell) headerRow.createCell(columnIndex);
        cell.setCellValue("category");
        rowIndex++;
        Cell firstDataCell = null;
        Cell nextCell = null;
        for(BulkTemplateDTO dto : bulkTemplateDTOs) {
            Row dataRow = dataSheet.getRow(rowIndex) != null ? dataSheet.getRow(rowIndex) : dataSheet.createRow(rowIndex);
            nextCell = dataRow.getCell(columnIndex) != null ? dataRow.getCell(columnIndex) : dataRow.createCell(columnIndex);
            nextCell.setCellValue(dto.getCategoryName());
            if(firstDataCell == null) {
                firstDataCell = nextCell;
            }
            rowIndex++;
        }
        columnIndex++;
        Name namedCell = dataSheet.getWorkbook().createName();
        namedCell.setNameName("category");
        namedCell.setRefersToFormula(DATA_SHEET + "!" + DOLLAR + firstDataCell.getAddress().toString().split("\\d+")[0] +  DOLLAR + firstDataCell.getAddress().toString().split("[a-zA-Z]+")[1] + ":" + DOLLAR + nextCell.getAddress().toString().split("\\d+")[0] +  DOLLAR + nextCell.getAddress().toString().split("[a-zA-Z]+")[1]);
        lookUpIndexInfoMap.put("category", Arrays.asList(firstDataCell.getAddress().toString(), nextCell.getAddress().toString()));
        return columnIndex;
    }

    private int populateDataSheetForFields(String categoryName, Sheet dataSheet, List<BulkTemplateDTO.ColumnDTO> columns, Map<String, List<String>> lookUpIndexInfoMap, String fieldType, int columnIndex) {
        for (BulkTemplateDTO.ColumnDTO column : columns) {
            int rowIndex = 0;
            if ("Dropdown".equalsIgnoreCase(column.getType())) {

                Row headerRow = dataSheet.getRow(rowIndex) != null ? dataSheet.getRow(rowIndex) : dataSheet.createRow(rowIndex);
                XSSFCell cell = (XSSFCell) headerRow.createCell(columnIndex);
                cell.setCellValue(fieldType);
                rowIndex++;

                headerRow = dataSheet.getRow(rowIndex) != null ? dataSheet.getRow(rowIndex) : dataSheet.createRow(rowIndex);
                cell = (XSSFCell) headerRow.createCell(columnIndex);
                cell.setCellValue(column.getColumnName());
                rowIndex++;
                Cell firstDataCell = null;
                Cell nextCell = null;
                for(String value : column.getAllowedValues()) {
                    Row dataRow = dataSheet.getRow(rowIndex) != null ? dataSheet.getRow(rowIndex) : dataSheet.createRow(rowIndex);
                    nextCell = dataRow.getCell(columnIndex) != null ? dataRow.getCell(columnIndex) : dataRow.createCell(columnIndex);
                    nextCell.setCellValue(value);
                    if(firstDataCell == null) {
                        firstDataCell = nextCell;
                    }
                    rowIndex++;
                }
                Name namedCell = dataSheet.getWorkbook().createName();
                namedCell.setNameName(categoryName.replaceAll(" ", "_") + column.getColumnName().replaceAll(" ", "_"));
                namedCell.setRefersToFormula(DATA_SHEET + "!" + DOLLAR + firstDataCell.getAddress().toString().split("\\d+")[0] +  DOLLAR + firstDataCell.getAddress().toString().split("[a-zA-Z]+")[1] + ":" + DOLLAR + nextCell.getAddress().toString().split("\\d+")[0] +  DOLLAR + nextCell.getAddress().toString().split("[a-zA-Z]+")[1]);
                columnIndex++;
                lookUpIndexInfoMap.put(column.getColumnName(), Arrays.asList(firstDataCell.getAddress().toString(), nextCell.getAddress().toString()));
            }
        }
        return columnIndex;
    }

    public static ByteArrayResource getByteArrayResource(Workbook workbook){
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try {
            workbook.write(bos);
        }
        catch(IOException e){
            log.error("Error generating the report excel", e);
        }
        finally {
            IOUtils.closeQuietly(bos);
            IOUtils.closeQuietly(workbook);
        }
        byte[] bytes = bos.toByteArray();
        return new ByteArrayResource(bytes);
    }

}
