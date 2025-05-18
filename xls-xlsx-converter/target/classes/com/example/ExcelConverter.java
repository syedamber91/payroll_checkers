package com.example;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class ExcelConverter {

    public static void convertXlsToXlsx(String inputPath, String outputPath) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputPath);
             Workbook oldWorkbook = new HSSFWorkbook(fis);
             Workbook newWorkbook = new XSSFWorkbook()) {

            // Map to track font conversions
            Map<Short, Font> fontMap = new HashMap<>();

            // Copy all sheets
            for (int i = 0; i < oldWorkbook.getNumberOfSheets(); i++) {
                Sheet oldSheet = oldWorkbook.getSheetAt(i);
                Sheet newSheet = newWorkbook.createSheet(oldSheet.getSheetName());

                // Copy column widths
                for (int col = 0; col < oldSheet.getRow(0).getLastCellNum(); col++) {
                    newSheet.setColumnWidth(col, oldSheet.getColumnWidth(col));
                }

                // Copy merged regions
                for (int j = 0; j < oldSheet.getNumMergedRegions(); j++) {
                    CellRangeAddress mergedRegion = oldSheet.getMergedRegion(j);
                    newSheet.addMergedRegion(mergedRegion);
                }

                // Copy all rows
                for (int rowIndex = 0; rowIndex <= oldSheet.getLastRowNum(); rowIndex++) {
                    Row oldRow = oldSheet.getRow(rowIndex);
                    if (oldRow == null) continue;

                    Row newRow = newSheet.createRow(rowIndex);
                    newRow.setHeight(oldRow.getHeight());

                    // Copy all cells
                    for (int colIndex = 0; colIndex < oldRow.getLastCellNum(); colIndex++) {
                        Cell oldCell = oldRow.getCell(colIndex);
                        if (oldCell == null) continue;

                        Cell newCell = newRow.createCell(colIndex);
                        copyCell(oldWorkbook, newWorkbook, oldCell, newCell, fontMap);
                    }
                }
            }

            // Save the new workbook
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                newWorkbook.write(fos);
            }
        }
    }

    private static void copyCell(Workbook oldWorkbook, Workbook newWorkbook, Cell oldCell, Cell newCell, Map<Short, Font> fontMap) {
        // Copy cell value
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                handleFormula(oldCell, newCell);
                break;
            case BLANK:
                newCell.setBlank();
                break;
            default:
                newCell.setBlank();
        }

        // Copy cell style
        CellStyle oldStyle = oldCell.getCellStyle();
        CellStyle newStyle = newWorkbook.createCellStyle();
        copyCellStyle(oldWorkbook, newWorkbook, oldStyle, newStyle, fontMap);
        newCell.setCellStyle(newStyle);
    }

    private static void copyCellStyle(Workbook oldWorkbook, Workbook newWorkbook, CellStyle oldStyle, CellStyle newStyle, Map<Short, Font> fontMap) {
        // Handle font
        Font oldFont = oldWorkbook.getFontAt(oldStyle.getFontIndex());
        Font newFont;
    
        short fontIndex = (short) oldFont.getIndex();  // Explicit cast to short
    
        if (fontMap.containsKey(fontIndex)) {
            newFont = fontMap.get(fontIndex);
        } else {
            newFont = newWorkbook.createFont();
            newFont.setBold(oldFont.getBold());
            newFont.setColor(oldFont.getColor());
            newFont.setFontHeightInPoints(oldFont.getFontHeightInPoints());
            newFont.setFontName(oldFont.getFontName());
            newFont.setItalic(oldFont.getItalic());
            newFont.setStrikeout(oldFont.getStrikeout());
            newFont.setTypeOffset(oldFont.getTypeOffset());
            newFont.setUnderline(oldFont.getUnderline());
            fontMap.put(fontIndex, newFont);
        }
        newStyle.setFont(newFont);
    
        // Copy other style properties
        newStyle.setAlignment(oldStyle.getAlignment());
        newStyle.setBorderBottom(oldStyle.getBorderBottom());
        newStyle.setBorderLeft(oldStyle.getBorderLeft());
        newStyle.setBorderRight(oldStyle.getBorderRight());
        newStyle.setBorderTop(oldStyle.getBorderTop());
        newStyle.setBottomBorderColor(oldStyle.getBottomBorderColor());
        newStyle.setDataFormat(oldStyle.getDataFormat());
        newStyle.setFillBackgroundColor(oldStyle.getFillBackgroundColor());
        newStyle.setFillForegroundColor(oldStyle.getFillForegroundColor());
        newStyle.setFillPattern(oldStyle.getFillPattern());
        newStyle.setHidden(oldStyle.getHidden());
        newStyle.setIndention(oldStyle.getIndention());
        newStyle.setLeftBorderColor(oldStyle.getLeftBorderColor());
        newStyle.setLocked(oldStyle.getLocked());
        newStyle.setRightBorderColor(oldStyle.getRightBorderColor());
        newStyle.setRotation(oldStyle.getRotation());
        newStyle.setTopBorderColor(oldStyle.getTopBorderColor());
        newStyle.setVerticalAlignment(oldStyle.getVerticalAlignment());
        newStyle.setWrapText(oldStyle.getWrapText());
    }

    private static void handleFormula(Cell oldCell, Cell newCell) {
        String formula = oldCell.getCellFormula();
        
        // Check for external references
        if (formula.contains(".xlsx]") || formula.contains(".xls]")) {
            System.out.println("Converting external reference formula: " + formula);
            
            // Convert to internal reference if possible
            String convertedFormula = formula
                .replaceAll("\\['?[^]]+\\.xlsx?\\]'?", "")
                .replaceAll("''", "'"); // Clean up any double quotes
            
            try {
                newCell.setCellFormula(convertedFormula);
                System.out.println("Converted to: " + convertedFormula);
            } catch (Exception e) {
                System.err.println("Failed to convert formula - using cached value");
                useCachedValue(oldCell, newCell);
            }
        } else {
            try {
                newCell.setCellFormula(formula);
            } catch (Exception e) {
                System.err.println("Formula conversion failed - using cached value");
                useCachedValue(oldCell, newCell);
            }
        }
    }

    private static void useCachedValue(Cell oldCell, Cell newCell) {
        switch (oldCell.getCachedFormulaResultType()) {
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            default:
                newCell.setBlank();
        }
    }

    public static void main(String[] args) {
        if (args.length < 2) {
            System.out.println("Usage: java -jar converter.jar input.xls output.xlsx");
            return;
        }

        try {
            convertXlsToXlsx(args[0], args[1]);
            System.out.println("Conversion completed successfully!");
        } catch (IOException e) {
            System.err.println("Conversion failed:");
            e.printStackTrace();
            System.exit(1);
        }
    }
}