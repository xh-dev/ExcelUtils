package me.xethh.util.excelUtils.reading;

import com.sun.corba.se.spi.orbutil.threadpool.Work;
import me.xethh.util.excelUtils.common.ExcelReadValue;
import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.util.excelUtils.model.CellStyleScanningModel;
import me.xethh.utils.wrapper.Tuple2;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class WorkbookScanning {
    public static class ScanningIterator implements Iterator<CellScanningModel> {
        private Workbook workbook;
        private Iterator<Integer> itSheetIndex;
        private Iterator<Row> rowIt;
        private Iterator<Cell> cellIt;
        private ScanningIterator(Workbook workbook){
            this.workbook = workbook;
            List<Integer> listSheet = new ArrayList<>();
            for(int i=0;i<workbook.getNumberOfSheets();i++)
                listSheet.add(i);
            itSheetIndex = listSheet.listIterator();
            if(null!=itSheetIndex && itSheetIndex.hasNext()){
                rowIt = workbook.getSheetAt(itSheetIndex.next()).iterator();
                if(null!=null && rowIt.hasNext()){
                    this.cellIt = rowIt.next().cellIterator();
                }
            }
        }
        @Override
        public boolean hasNext() {
            while(true){
                if(null!=cellIt && cellIt.hasNext()){
                    return true;
                }
                else{
                    if(null!=rowIt && rowIt.hasNext()){
                        cellIt = rowIt.next().iterator();
                        continue;
                    }
                    else{
                        if(null!=itSheetIndex && itSheetIndex.hasNext()){
                            rowIt = workbook.getSheetAt(itSheetIndex.next()).iterator();
                            continue;
                        }
                        else{
                            return false;
                        }
                    }
                }
            }
        }

        @Override
        public CellScanningModel next() {
            Cell cell = cellIt.next();
            CellScanningModel model = new CellScanningModel();
            model.setSheetName(cell.getSheet().getSheetName());
            model.setActRow(cell.getRowIndex());
            model.setActCol(cell.getColumnIndex());
            model.setCellStyle(new CellStyleScanningModel(workbook,cell.getCellStyle()));
            model.setCellStr("");
            Tuple2<CellScanningModel.CellType, Object> value = ExcelReadValue.read(cell);
            model.setValue(value.getV2());
            model.setCellType(value.getV1());
            return model;
        }

        @Override
        public void remove() {

        }
    }
    public static Iterator<CellScanningModel> scan(Workbook workbook){
        return new ScanningIterator(workbook);
    }

    public static Iterator<String[]> scanAsArr(Workbook workbook){
        final Iterator<CellScanningModel> it = scan(workbook);
        return new Iterator<String[]>() {
            boolean first = true;
            @Override
            public boolean hasNext() {
                if(first) return true;
                return it.hasNext();
            }

            private String colorToString(short[] argb){
                if(argb==null) return "";
                String tArgb  = "";
                for(int i=0; i < argb.length;i++){
                    tArgb += String.format("%d", argb[i]);
                    if(i<(argb.length-1)){
                        tArgb+=",";
                    }
                }
                return tArgb;
            }

            @Override
            public String[] next() {
                String[] arr = new String[31];
                if(first){
                    first=false;
                    arr[0] = "Sheet Name";
                    arr[1] = "Row index";
                    arr[2] = "Col index";
                    arr[3] = "Cell String";
                    arr[4] = "Cell Type";
                    arr[5] = "Cell Value";
                    arr[6] = "Data Format";
                    arr[7] = "Background Color";

                    arr[8] = "Background ARGB";
                    arr[9] = "Border bot";
                    arr[10] = "Border bot color";
                    arr[11] = "Border top";
                    arr[12] = "Border top color";
                    arr[13] = "Border left";
                    arr[14] = "Border left color";
                    arr[15] = "Border right";
                    arr[16] = "Border right color";
                    arr[17] = "Fill Pattern";
                    arr[18] = "Foreground color";
                    arr[19] = "Foreground ARGB";
                    arr[20] = "Horizontal Alignment";
                    arr[21] = "Indentation";
                    arr[22] = "Rotation";
                    arr[23] = "Vertical Alignment";
                    arr[24] = "Font Name";
                    arr[25] = "Char set";
                    arr[26] = "Font color";
                    arr[27] = "Font color RGB";
                    arr[28] = "Font Height";
                    arr[29] = "Under line";
                    arr[30] = "Type offset";
                    return arr;
                }
                CellScanningModel cellModel = it.next();
                arr[0] = cellModel.getSheetName();
                arr[1] = String.valueOf(cellModel.getActRow());
                arr[2] = String.valueOf(cellModel.getActCol());
                arr[3] = cellModel.getCellStr();
                arr[4] = cellModel.getCellType().name();
                arr[5] = cellModel.getValue()==null?"[null]":cellModel.getValue().toString();
                CellStyleScanningModel style = cellModel.getCellStyle();
                arr[6] = style.getDataFromat();
                arr[7] = style.getBackgroundColor().name();

                arr[8] = colorToString(style.getBackgroundRGB());
                arr[9] = style.getBorderBot().name();
                arr[10] = style.getBorderBotColor().name();
                arr[11] = style.getBorderTop().name();
                arr[12] = style.getBorderTopColor().name();
                arr[13] = style.getBorderLeft().name();
                arr[14] = style.getBorderLeftColor().name();
                arr[15] = style.getBorderRight().name();
                arr[16] = style.getBorderRightColor().name();
                arr[17] = style.getFillPatternType().name();
                arr[18] = style.getForegroundColor().name();
                arr[19] = colorToString(style.getForegroundRGB());
                arr[20] = style.getHorizontalAlignment().name();
                arr[21] = style.getIndentation()+"";
                arr[22] = style.getRotation()+"";
                arr[23] = style.getVerticalAlignment().name();
                CellStyleScanningModel.Font font = style.getFont();
                arr[24] = font.getName();
                arr[25] = font.getCharSet()+"";
                arr[26] = font.getColor()+"";
                arr[27] = colorToString(font.getColorRgb());
                arr[28] = font.getFontHeightInPoint()+"";
                arr[29] = font.getIsUnderLine()+"";
                arr[30] = font.getTypeOffset()+"";

                return arr;
            }

            @Override
            public void remove() {

            }
        };
    }
}
