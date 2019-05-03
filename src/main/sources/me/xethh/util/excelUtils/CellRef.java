package me.xethh.util.excelUtils;

import org.apache.poi.ss.usermodel.CellType;

public class CellRef {
    private Integer row;
    private Integer col;
    private String colStr;
    private CellType cellType;
    private Object value;

    public Integer getRow() {
        return row;
    }

    public void setRow(Integer row) {
        this.row = row;
    }

    public Integer getCol() {
        return col;
    }

    public void setCol(Integer col) {
        this.col = col;
    }

    public String getColStr() {
        return colStr;
    }

    public void setColStr(String colStr) {
        this.colStr = colStr;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public CellType getCellType() {
        return cellType;
    }

    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }
}
