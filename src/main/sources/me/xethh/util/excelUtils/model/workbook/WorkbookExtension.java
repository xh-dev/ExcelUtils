package me.xethh.util.excelUtils.model.workbook;

import me.xethh.util.excelUtils.model.sheet.SheetExtension;
import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookExtension extends Workbook {
    SheetExtension createSheet();
    SheetExtension createSheet(String s);
    SheetExtension cloneSheet(int i);
    SheetExtension getSheetAt(int i);
    SheetExtension getSheet(String s);
    static WorkbookExtension extendsWorkbook(Workbook workbook){
        return new WorkbookExtensionImpl(workbook);
    }
}
