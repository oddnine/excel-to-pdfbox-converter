package com.example.excel2pdf.service;

import lombok.RequiredArgsConstructor;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

@Service
@RequiredArgsConstructor
public class ExcelToPDFConverterService {

    public void convertExcelToPDF(File excelFile) throws IOException {
        try (FileInputStream excelFileStream = new FileInputStream(excelFile);
             Workbook workbook = new XSSFWorkbook(excelFileStream);
             PDDocument pdfDocument = new PDDocument()) {

            PDType0Font customFont = PDType0Font.load(pdfDocument, new File(new ClassPathResource("fonts/NanumGothic.ttf").getURI()));
            PDType0Font customFontBold = PDType0Font.load(pdfDocument, new File(new ClassPathResource("fonts/NanumGothicBold.ttf").getURI()));

            for (int sheetIndex = 0; sheetIndex < Math.min(workbook.getNumberOfSheets(), 2); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                PDPage page = new PDPage(PDRectangle.A4);
                pdfDocument.addPage(page);

                Map<CellAddress, CellRangeAddress> mergedCellsMap = getMergedCellsMap(sheet);

                try (PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page)) {
                    float yPosition = PDRectangle.A4.getHeight() - 20;

                    for (Row row : sheet) {
                        float xPosition = 20;
                        float cellHeight = row.getHeightInPoints();

                        for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                            Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            CellAddress cellAddress = cell.getAddress();

                            if (mergedCellsMap.containsKey(cellAddress)) {
                                CellRangeAddress cellRange = mergedCellsMap.get(cellAddress);

                                if (cellRange.getFirstRow() == cell.getRowIndex() && cellRange.getFirstColumn() == cell.getColumnIndex()) {
                                    float cellWidth = 0;
                                    for (int col = cellRange.getFirstColumn(); col <= cellRange.getLastColumn(); col++) {
                                        cellWidth += getExcelCellWidthInPoints(sheet, col);
                                    }
                                    cellHeight = getMergedCellHeight(sheet, cellRange);

                                    drawMergedCellContent(contentStream, sheet, cellRange, cell, xPosition, yPosition, cellWidth, cellHeight, customFont, customFontBold, workbook);

                                    cellIndex = cellRange.getLastColumn();
                                    xPosition += cellWidth;
                                }
                            } else {
                                float cellWidth = getExcelCellWidthInPoints(sheet, cellIndex);

                                drawCellContent(contentStream, cell, xPosition, yPosition, cellWidth, cellHeight, customFont, customFontBold, workbook);

                                xPosition += cellWidth;
                            }
                        }
                        yPosition -= cellHeight;
                    }
                }
            }
            pdfDocument.save("output.pdf");
        }
    }

    private void drawMergedCellContent(PDPageContentStream contentStream, Sheet sheet, CellRangeAddress cellRange, Cell cell, float xPosition, float yPosition, float cellWidth, float cellHeight, PDType0Font customFont, PDType0Font customBoldFont, Workbook workbook) throws IOException {
        CellStyle cellStyle = cell.getCellStyle();
        Font cellFont = workbook.getFontAt(cellStyle.getFontIndex());
        float fontSize = cellFont.getFontHeightInPoints();

        Color bgColor = getExcelCellBackgroundColor(cellStyle);
        if (bgColor != null) {
            contentStream.setNonStrokingColor(bgColor);
            contentStream.addRect(xPosition, yPosition - cellHeight, cellWidth, cellHeight);
            contentStream.fill();
        }

        contentStream.setNonStrokingColor(Color.BLACK);

        if (cellFont.getBold()) {
            contentStream.setFont(customFont, fontSize);
        } else {
            contentStream.setFont(customBoldFont, fontSize);
        }

        String text = getStringCellValue(cell);
        String[] lines = text.split("\n");
        float totalTextHeight = lines.length * fontSize * 1.2f;

        float verticalOffset = calculateVerticalOffset(cellStyle, cellHeight, totalTextHeight, fontSize);

        float currentYPosition = yPosition - fontSize - verticalOffset;
        for (String line : lines) {
            float adjustedXPosition = calculateAdjustedXPosition(xPosition, cellStyle, customFont, fontSize, line, cellWidth);
            contentStream.beginText();
            contentStream.newLineAtOffset(adjustedXPosition, currentYPosition);
            contentStream.showText(line);
            contentStream.endText();
            currentYPosition -= fontSize * 1.2f;
        }

        drawCellBorders(contentStream, cellStyle, xPosition, yPosition, cellWidth, cellHeight);
    }

    private void drawCellContent(PDPageContentStream contentStream, Cell cell, float xPosition, float yPosition, float cellWidth, float cellHeight, PDType0Font customFont, PDType0Font customFontBold, Workbook workbook) throws IOException {
        CellStyle cellStyle = cell.getCellStyle();
        Font cellFont = workbook.getFontAt(cellStyle.getFontIndex());
        float fontSize = cellFont.getFontHeightInPoints();

        Color bgColor = getExcelCellBackgroundColor(cellStyle);
        if (bgColor != null) {
            contentStream.setNonStrokingColor(bgColor);
            contentStream.addRect(xPosition, yPosition - cellHeight, cellWidth, cellHeight);
            contentStream.fill();
        }

        contentStream.setNonStrokingColor(Color.BLACK);

        if (cellFont.getBold()) {
            contentStream.setFont(customFont, fontSize);
        } else {
            contentStream.setFont(customFontBold, fontSize);
        }

        String text = getStringCellValue(cell);
        String[] lines = text.split("\n");
        float totalTextHeight = lines.length * fontSize * 1.2f;

        float verticalOffset = calculateVerticalOffset(cellStyle, cellHeight, totalTextHeight, fontSize);

        float currentYPosition = yPosition - fontSize - verticalOffset;
        for (String line : lines) {
            float adjustedXPosition = calculateAdjustedXPosition(xPosition, cellStyle, customFont, fontSize, line, cellWidth);
            contentStream.beginText();
            contentStream.newLineAtOffset(adjustedXPosition, currentYPosition);
            contentStream.showText(line);
            contentStream.endText();
            currentYPosition -= fontSize * 1.2f;
        }

        drawCellBorders(contentStream, cellStyle, xPosition, yPosition, cellWidth, cellHeight);
    }

    private Map<CellAddress, CellRangeAddress> getMergedCellsMap(Sheet sheet) {
        Map<CellAddress, CellRangeAddress> mergedCellsMap = new HashMap<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRange = sheet.getMergedRegion(i);
            for (int row = cellRange.getFirstRow(); row <= cellRange.getLastRow(); row++) {
                for (int col = cellRange.getFirstColumn(); col <= cellRange.getLastColumn(); col++) {
                    mergedCellsMap.put(new CellAddress(row, col), cellRange);
                }
            }
        }
        return mergedCellsMap;
    }

    private float getMergedCellHeight(Sheet sheet, CellRangeAddress cellRange) {
        float height = 0;
        for (int row = cellRange.getFirstRow(); row <= cellRange.getLastRow(); row++) {
            height += sheet.getRow(row).getHeightInPoints();
        }
        return height;
    }

    private Color getExcelCellBackgroundColor(CellStyle cellStyle) {
        if (cellStyle.getFillPattern() == FillPatternType.SOLID_FOREGROUND) {
            XSSFColor color = (XSSFColor) cellStyle.getFillForegroundColorColor();
            if (color != null) {
                byte[] rgb = color.getRGB();
                if (rgb != null) {
                    return new Color((rgb[0] & 0xFF), (rgb[1] & 0xFF), (rgb[2] & 0xFF));
                }
            }
        }
        return null;
    }

    private float calculateAdjustedXPosition(float xPosition, CellStyle cellStyle, PDType0Font customFont, float fontSize, String text, float cellWidth) throws IOException {
        float textWidth = customFont.getStringWidth(text) / 1000 * fontSize;
        HorizontalAlignment alignment = cellStyle.getAlignment();

        switch (alignment) {
            case CENTER:
                return xPosition + (cellWidth - textWidth) / 2;
            case RIGHT:
                return xPosition + cellWidth - textWidth;
            case LEFT:
            default:
                return xPosition;
        }
    }

    private float calculateVerticalOffset(CellStyle cellStyle, float cellHeight, float totalTextHeight, float fontSize) {
        VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();
        switch (verticalAlignment) {
            case CENTER:
                return (cellHeight - totalTextHeight) / 2;
            case TOP:
                return 0;
            case BOTTOM:
            default:
                return cellHeight - totalTextHeight - fontSize * 0.2f;
        }
    }

    private void drawCellBorders(PDPageContentStream contentStream, CellStyle style, float xPosition, float yPosition, float cellWidth, float cellHeight) throws IOException {

        if (style.getBorderTop() != BorderStyle.NONE) {
            contentStream.moveTo(xPosition, yPosition);
            contentStream.lineTo(xPosition + cellWidth, yPosition);
            contentStream.stroke();
        }

        if (style.getBorderBottom() != BorderStyle.NONE) {
            contentStream.moveTo(xPosition, yPosition - cellHeight);
            contentStream.lineTo(xPosition + cellWidth, yPosition - cellHeight);
            contentStream.stroke();
        }

        if (style.getBorderLeft() != BorderStyle.NONE) {
            contentStream.moveTo(xPosition, yPosition);
            contentStream.lineTo(xPosition, yPosition - cellHeight);
            contentStream.stroke();
        }

        if (style.getBorderRight() != BorderStyle.NONE) {
            contentStream.moveTo(xPosition + cellWidth, yPosition);
            contentStream.lineTo(xPosition + cellWidth, yPosition - cellHeight);
            contentStream.stroke();
        }
    }

    private String getStringCellValue(Cell cell) {
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    private float getExcelCellWidthInPoints(Sheet sheet, int columnIndex) {
        int widthUnits = sheet.getColumnWidth(columnIndex);
        return widthUnits * 7.00174f / 256;
    }
}
