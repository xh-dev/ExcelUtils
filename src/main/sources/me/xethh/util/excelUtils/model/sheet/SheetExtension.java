package me.xethh.util.excelUtils.model.sheet;

import me.xethh.util.excelUtils.model.col.CellExtension;
import me.xethh.util.excelUtils.model.row.RowExtension;
import me.xethh.util.excelUtils.model.workbook.WorkbookExtension;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Iterator;
import java.util.List;
import java.util.Spliterators;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public interface SheetExtension extends Sheet {
    RowExtension createRow(int i);
    RowExtension getRow(int i);
    WorkbookExtension getWorkbook();
    static SheetExtension extendsSheet(Sheet sheet){
        return new SheetExtensionImpl(sheet);
    }

    Iterator<CellExtension> column(int col);
    default Stream<CellExtension> columnAsStream(int col){
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(this.column(col), 0), false);
    }
    default List<CellExtension> columnAsList(int col){
        return columnAsStream(col).collect(Collectors.toList());
    }
}
