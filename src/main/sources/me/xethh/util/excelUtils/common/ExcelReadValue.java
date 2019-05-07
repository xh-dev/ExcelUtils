package me.xethh.util.excelUtils.common;

import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.utils.wrapper.Tuple2;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;

public class ExcelReadValue {
    public static int[] toIntArr(String[] ss){
        int[] is = new int[ss.length];
        for(int i=0;i<ss.length;i++){
            is[i] = Integer.parseInt(ss[i]);
        }
        return is;
    }
    public static byte[] toByteArr(String[] ss){
        int[] is = toIntArr(ss);
        byte[] ba = new byte[ss.length];
        for(int i=0;i<is.length;i++){
            ba[i] = (byte) is[i];
        }
        return ba;
    }
    public static boolean isPureDark(String[] ss){
        int[] s = toIntArr(ss);
        if(s.length==3){
            return (s[0]+s[1]+s[2])==0;
        }
        else if(s.length==4){
            return (s[1]+s[2]+s[3])==0;
        }
        throw new RuntimeException("Array not supported: "+ss);
    }
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
