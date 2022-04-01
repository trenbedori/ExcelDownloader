package utils;

import exception.ImpossibleCellWrite;
import exception.ImpossibleFieldAccess;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import resource.data.ExcelData;
import resource.data.ExcelDataFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class ExcelDownloader {

    public void excelDownload(List<?> data, Class<?> clazzType, HttpServletResponse response, String fileName, boolean useSeq) {
        final Workbook workbook = new SXSSFWorkbook();
        final Sheet sheet = workbook.createSheet("Sheet");
        final ExcelData excelData = ExcelDataFactory.getExcelData(clazzType, workbook);

        setHttpHeader(response, fileName);

        renderHeader(sheet, excelData, clazzType, useSeq);
        renderBody(sheet, excelData, data, clazzType, useSeq);

        write(workbook, getOutputStream(response));
    }

    public void excelDownload(List<?> data, Class<?> clazzType, FileOutputStream stream, boolean useSeq) {
        final Workbook workbook = new SXSSFWorkbook();
        final Sheet sheet = workbook.createSheet("Sheet");
        final ExcelData excelData = ExcelDataFactory.getExcelData(clazzType, workbook);

        renderHeader(sheet, excelData, clazzType, useSeq);
        renderBody(sheet, excelData, data, clazzType, useSeq);

        write(workbook, stream);
    }

    private void setHttpHeader(HttpServletResponse response, String fileName) {
        response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".xlsx");
        response.setContentType("application/download;charset=utf-8");
        response.setHeader("Content-Transfer-Encoding", "binary");
    }

    private void renderHeader(Sheet sheet, ExcelData excelData, Class<?> clazzType, boolean useSeq) {
        Row row = sheet.createRow(0);

        int col = 0;
        for(Field field : clazzType.getDeclaredFields()) {
            if(field.isAnnotationPresent(annotation.Cell.class)) {
                if (col == 0 && useSeq) {
                    Cell cell = row.createCell(col++);
                    setCellStyle(cell, excelData.getDefaultHeaderStyle());
                    setCellValue(cell, "No");
                }
                String headerName = excelData.getHeaderMap().get(getFieldName(field));

                Cell cell = row.createCell(col++);
                setCellStyle(cell, excelData.getHeaderStyleMap().get(getFieldName(field)));
                setCellValue(cell, headerName);
            }
        }
    }

    private void renderBody(Sheet sheet, ExcelData excelData, List<?> data, Class<?> clazzType, boolean useSeq) {
        int ROW_INDEX = 1;

        for(Object object : data) {
            Row row = sheet.createRow(ROW_INDEX++);

            int COL_INDEX = 0;
            for(String fieldName : excelData.getFieldName()) {
                if(COL_INDEX==0 && useSeq) {
                    Cell cell = row.createCell(COL_INDEX++);
                    setCellStyle(cell, excelData.getDefaultBodyStyle());
                    setCellValue(cell, Integer.toString(ROW_INDEX-1));
                }

                Field field = getField(clazzType, fieldName);
                field.setAccessible(true);

                Cell cell = row.createCell(COL_INDEX++);
                setCellStyle(cell, excelData.getBodyStyleMap().get(getFieldName(field)));
                setCellValue(cell, cellObjectToString(getFieldObject(field, object)));
            }
        }
    }

    private void setCellStyle(Cell cell, CellStyle cellStyle) {
        cell.setCellStyle(cellStyle);
    }

    private void setCellValue(Cell cell, String value) {
        cell.setCellValue(value);
    }

    private Field getField(Class<?> clazz, String fieldName) {
        for(Field field : clazz.getDeclaredFields()) {
            if(field.getName().equals(fieldName)) {
                return field;
            }
        }
        return null;
    }

    private String getFieldName(Field field) {
        return field.getName();
    }

    private Object getFieldObject(Field field, Object object) {
        try {
            return field.get(object);
        } catch (IllegalAccessException e) {
            throw new ImpossibleFieldAccess(e.getMessage());
        }
    }

    private String cellObjectToString(Object cellValue) {
        if (cellValue == null) {
            return "";
        } else if(cellValue instanceof LocalDateTime) {
            return ((LocalDateTime) cellValue).format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        } else {
            return cellValue.toString();
        }
    }

    private OutputStream getOutputStream(HttpServletResponse response) {
        try { return response.getOutputStream(); }
        catch (IOException e) { throw new ImpossibleCellWrite(e.getMessage()); }
    }

    private void write(Workbook workbook, OutputStream outputStream) {
        try { workbook.write(outputStream); } catch (IOException e) { throw new ImpossibleCellWrite(e.getMessage()); }
        finally { try { outputStream.close(); } catch (IOException e) { throw new ImpossibleCellWrite(e.getMessage()); } }
    }
}
