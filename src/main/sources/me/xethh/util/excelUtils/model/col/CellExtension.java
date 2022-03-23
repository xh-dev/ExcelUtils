package me.xethh.util.excelUtils.model.col;


import me.xethh.util.excelUtils.model.row.RowExtension;
import me.xethh.util.excelUtils.model.sheet.SheetExtension;
import org.apache.poi.ss.usermodel.Cell;

public interface CellExtension extends Cell {
    SheetExtension getSheet();
    RowExtension getRow();
    static CellExtension extendsCell(Cell cell){
        return new CellExtensionImpl(cell);
    }
}
