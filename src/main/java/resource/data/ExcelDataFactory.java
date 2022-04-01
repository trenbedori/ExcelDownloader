package resource.data;

import annotation.AllCell;
import annotation.Cell;
import annotation.CustomCellStyle;
import exception.ImpossibleCallCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import resource.style.ExcelCellStyle;
import resource.style.NoExcelCellStyle;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelDataFactory {
    private static final int HEADER = 0;
    private static final int BODY = 1;

    public static ExcelData getExcelData(Class<?> clazz, Workbook wb) {
        List<String> fieldName = new ArrayList<>();
        Map<String, String> headerMap = new HashMap<>();
        Map<String, CellStyle> headerStyleMap = new HashMap<>();
        Map<String, CellStyle> bodyStyleMap = new HashMap<>();
        CellStyle defaultHeaderStyle = getDefaultCellStyle(clazz, wb, HEADER);
        CellStyle defaultBodyStyle = getDefaultCellStyle(clazz, wb, BODY);

        for(Field field : clazz.getDeclaredFields()) {
            if(field.isAnnotationPresent(Cell.class)) {
                Cell cell = field.getAnnotation(Cell.class);

                fieldName.add(field.getName());
                headerMap.put(field.getName(), cell.headerName());

                CustomCellStyle headerCellStyle = cell.headerStyle();
                CustomCellStyle bodyCellStyle = cell.bodyStyle();

                headerStyleMap.put(field.getName(), getCellStyle(clazz, headerCellStyle, wb, HEADER));
                bodyStyleMap.put(field.getName(), getCellStyle(clazz, bodyCellStyle, wb, BODY));
            }
        }

        return new ExcelData(fieldName, headerMap, headerStyleMap, bodyStyleMap, defaultHeaderStyle, defaultBodyStyle);
    }

    private static CellStyle getCellStyle(Class<?> clazz, CustomCellStyle customCellStyle, Workbook wb, int location) {
        CellStyle cellStyle = wb.createCellStyle();

        if(isCustom(customCellStyle)) {
            ExcelCellStyle excelCellStyle = getCellStyleInstance(customCellStyle.excelCellStyle());
            excelCellStyle.styleApply(cellStyle);
        } else {
            AllCell allCell = clazz.getAnnotation(AllCell.class);
            ExcelCellStyle excelCellStyle = getCellStyleInstance(location == HEADER ? allCell.headerStyle().excelCellStyle() : allCell.bodyStyle().excelCellStyle());
            excelCellStyle.styleApply(cellStyle);
        }

        return cellStyle;
    }

    private static CellStyle getDefaultCellStyle(Class<?> clazz, Workbook wb, int location) {
        CellStyle cellStyle = wb.createCellStyle();

        AllCell allCell = clazz.getAnnotation(AllCell.class);
        ExcelCellStyle excelCellStyle;

        if (location == HEADER) excelCellStyle = getCellStyleInstance(allCell.headerStyle().excelCellStyle());
        else excelCellStyle = getCellStyleInstance(allCell.bodyStyle().excelCellStyle());

        excelCellStyle.styleApply(cellStyle);

        return cellStyle;
    }

    private static ExcelCellStyle getCellStyleInstance(Class<? extends ExcelCellStyle> cellClazz) {
        try {
            return cellClazz.getDeclaredConstructor().newInstance();
        } catch ( NoSuchMethodException | InvocationTargetException | InstantiationException | IllegalAccessException e){
            throw new ImpossibleCallCellStyle(e.getMessage());
        }
    }

    private static boolean isCustom(CustomCellStyle customCellStyle) {
        return !customCellStyle.excelCellStyle().equals(NoExcelCellStyle.class);
    }
}
