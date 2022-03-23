package me.xethh.util.excelUtils.model.sheet;

import me.xethh.util.excelUtils.model.col.CellExtension;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Iterator;

public class SheetExtensionImpl extends SheetExtensionAbstract {
    public SheetExtensionImpl(Sheet sheet) {
        super(sheet);
    }

    @Override
    public Iterator<CellExtension> column(int col) {

        return new Iterator<CellExtension>() {
            int curRow;
            final int lastRow = getLastRowNum();

            @Override
            public boolean hasNext() {
                while (true) {
                    if (curRow > lastRow) {
                        return false;
                    } else if (curRow == lastRow && getRow(curRow) != null && getRow(curRow).getCell(col) != null) {
                        return true;
                    } else if (getRow(curRow) != null) {
                        if (getRow(curRow).getCell(col) != null) {
                            return true;
                        }
                    }
                    curRow++;
                }
            }

            @Override
            public CellExtension next() {
                return getRow(curRow++).getCell(col);
            }
        };
    }
}
