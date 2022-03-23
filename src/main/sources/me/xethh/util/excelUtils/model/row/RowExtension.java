package me.xethh.util.excelUtils.model.row;

import me.xethh.util.excelUtils.model.col.CellExtension;
import me.xethh.util.excelUtils.model.sheet.SheetExtension;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public interface RowExtension extends Row {
    CellExtension createCell(int i);
    CellExtension createCell(int i, CellType cellType);
    CellExtension getCell(int i);
    CellExtension getCell(int i, MissingCellPolicy missingCellPolicy);
    SheetExtension getSheet();
    static RowExtension extendsRow(Row row){
        return new RowExtensionImpl(row);
    }
}
