package resource.style;

import org.apache.poi.ss.usermodel.CellStyle;

public class NoExcelCellStyle implements ExcelCellStyle {

    @Override
    public void styleApply(CellStyle cellStyle) {
        // No Style
    }
}
