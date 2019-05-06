package me.xethh.util.excelUtils.model;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellReference;

public class CellScanningModel {
    public enum CellType{
        Integer, Decimal, String, Blank, Error, Date, Boolean, Formula
    }
    private String sheetName;
    private Integer actRow;
    private Integer actCol;
    private CellType cellType;
    private Object value;
    private CellStyleScanningModel cellStyle;
    private String cellStr;


    public Integer getActRow() {
        return actRow;
    }

    public void setActRow(Integer actRow) {
        this.actRow = actRow;
    }

    public Integer getActCol() {
        return actCol;
    }

    public void setActCol(Integer actCol) {
        this.actCol = actCol;
    }

    public CellType getCellType() {
        return cellType;
    }

    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public CellStyleScanningModel getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyleScanningModel cellStyle) {
        this.cellStyle = cellStyle;
    }

    public String getCellStr() {
        if(this.cellStr==null) this.cellStr = CellReference.convertNumToColString(actCol)+(actRow+1);
        return this.cellStr;
    }

    public void setCellStr(String cellStr) {
        this.cellStr = getCellStr();
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    @Override
    public String toString() {
        return "CellScanningModel{" +
                "sheetName='" + sheetName + '\'' +
                ", actRow=" + actRow +
                ", actCol=" + actCol +
                ", cellType=" + cellType +
                ", value=" + value +
                ", cellStyle=" + cellStyle +
                ", cellStr='" + cellStr + '\'' +
                '}';
    }

}
