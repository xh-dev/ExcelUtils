package me.xethh.util.excelUtils.common;

import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.utils.wrapper.Tuple2;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;

public class ExcelReadValue {
    public static Tuple2<CellScanningModel.CellType, Object> read(Cell cell){
        switch (cell.getCellType()){
            case BLANK:
                return Tuple2.of(CellScanningModel.CellType.Blank,null);
            case ERROR:
                return Tuple2.of(CellScanningModel.CellType.Error,null);
            case STRING:
                return Tuple2.of(CellScanningModel.CellType.String,(Object)cell.getStringCellValue());
            case NUMERIC:
                if(HSSFDateUtil.isCellDateFormatted(cell))
                    return Tuple2.of(CellScanningModel.CellType.Date,(Object)cell.getDateCellValue());
                else
                    return Tuple2.of(CellScanningModel.CellType.Decimal,(Object) cell.getNumericCellValue());
            case BOOLEAN:
                return Tuple2.of(CellScanningModel.CellType.Boolean,(Object) cell.getBooleanCellValue());
            case FORMULA:
                return Tuple2.of(CellScanningModel.CellType.Formula,(Object) cell.getCellFormula());
        }
        throw new RuntimeException(String.format("Error while reading cell[%s%s]", CellReference.convertNumToColString(cell.getColumnIndex()),cell.getRowIndex()+1));
    }
}
