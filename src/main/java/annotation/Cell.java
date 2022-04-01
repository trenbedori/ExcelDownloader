package annotation;

import resource.style.NoExcelCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface Cell {
    String headerName() default "";
    CustomCellStyle headerStyle() default @CustomCellStyle(excelCellStyle = NoExcelCellStyle.class);
    CustomCellStyle bodyStyle() default @CustomCellStyle(excelCellStyle = NoExcelCellStyle.class);
}
