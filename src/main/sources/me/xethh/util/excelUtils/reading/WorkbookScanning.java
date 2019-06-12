package me.xethh.util.excelUtils.reading;

import me.xethh.util.excelUtils.common.ExcelReadValue;
import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.util.excelUtils.model.CellStyleScanningModel;
import me.xethh.utils.wrapper.Tuple2;
import me.xethh.utils.wrapper.Tuple3;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WorkbookScanning {
    public static class SheetScanningIterator implements Iterator<CellScanningModel> {
        private Workbook workbook;
        private Iterator<Tuple3<String,String,String>> iterator;
        private int startRow, startCol, endCol, endRow, currentRow, currentCol;
        private Sheet currentSheet;
        private Row currentExcelRow;
        private SheetScanningIterator(Workbook workbook, List<Tuple3<String, String, String>> areaList){
            this.workbook = workbook;
            iterator = areaList.iterator();
            if(null!=iterator && iterator.hasNext())
                setupStartEndPont();
        }
        private void setupStartEndPont(){
            Tuple3<String, String, String> range = iterator.next();
            currentSheet = workbook.getSheetAt(workbook.getSheetIndex(range.getV1()));
            Pattern pattern = Pattern.compile("([a-zA-Z]+)([0-9]+)");
            Matcher matcher = pattern.matcher(range.getV2());
            if(matcher.matches()){
                startRow = Integer.parseInt(matcher.group(2))-1;
                startCol = CellReference.convertColStringToIndex(matcher.group(1));
                currentRow = startRow;
                currentCol=startCol-1;
                currentExcelRow = null;
            }
            else {
                throw new RuntimeException(String.format("Unexpected error pattern not match %s[%s]",pattern.toString(),range.getV2()));
            }
            matcher = pattern.matcher(range.getV3());
            if(matcher.matches()){
                endRow = Integer.parseInt(matcher.group(2))-1;
                endCol = CellReference.convertColStringToIndex(matcher.group(1));
            }
            else {
                throw new RuntimeException(String.format("Unexpected error pattern not match %s[%s]",pattern.toString(),range.getV3()));
            }
        }
        @Override
        public boolean hasNext() {
            try {
                if (currentCol < endCol) {
                    currentCol++;
                    return currentSheet.getRow(currentRow).getCell(currentCol) != null || hasNext();
                }
                if (currentCol == endCol && currentRow < endRow) {
                    currentCol = startCol;
                    while (currentSheet.getRow(++currentRow)==null)
                        if(currentRow==endRow)
                            return false;

                    return currentSheet.getRow(currentRow).getCell(currentCol) != null || hasNext();
                }
                if (endRow == currentRow && endCol == currentCol) {
                    if (iterator.hasNext()) {
                        setupStartEndPont();
                        return currentSheet.getRow(currentRow).getCell(currentCol) != null || hasNext();
                    } else
                        return false;
                }
            }
            catch (Exception ex){
                ex.printStackTrace();
            }
            throw new RuntimeException("Unexpected error");
        }


        @Override
        public CellScanningModel next() {
            Cell cell = currentSheet.getRow(currentRow).getCell(currentCol);
            if(cell==null)
                return null;
            CellScanningModel model = new CellScanningModel();
            model.setSheetName(currentSheet.getSheetName());
            model.setActRow(currentRow);
            model.setActCol(currentCol);
            model.setCellStyle(new CellStyleScanningModel(workbook,cell.getCellStyle()));
            Tuple2<CellScanningModel.CellType, Object> value = ExcelReadValue.read(cell);
            model.setValue(value.getV2());
            model.setCellType(value.getV1());
            return model;
        }

        @Override
        public void remove() {

        }
    }
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
    public static Iterator<CellScanningModel> scan(Workbook workbook, List<Tuple3<String,String,String>> areaList){
        return new SheetScanningIterator(workbook, areaList);
    }
    public static Iterator<CellScanningModel> scan(Workbook workbook){
        return new ScanningIterator(workbook);
    }

    public static Iterator<String[]> scanAsArr(Workbook workbook, List<Tuple3<String,String,String>> areaList){
        return internalScanAsArr(scan(workbook, areaList));
    }
    public static Iterator<String[]> scanAsArr(Workbook workbook){
        return internalScanAsArr(scan(workbook));
    }
    private static Iterator<String[]> internalScanAsArr(final Iterator<CellScanningModel> it){
        return new Iterator<String[]>() {
            boolean first = true;
            @Override
            public boolean hasNext() {
                if(first) return true;
                return it.hasNext();
            }


            @Override
            public String[] next() {
                if(!first)
                    return it.next().toStringArr();

                first=false;
                return CellScanningModel.toStringArrHeader();
            }

            @Override
            public void remove() {

            }
        };
    }
}
