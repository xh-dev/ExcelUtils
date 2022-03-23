package me.xethh.util.excelUtils.model.row;

import me.xethh.util.excelUtils.model.col.CellExtension;
import me.xethh.util.excelUtils.model.sheet.SheetExtension;
import org.apache.poi.ss.usermodel.*;

import java.util.Iterator;

public class RowExtensionAbstract implements RowExtension{
    protected final Row row;
    protected RowExtensionAbstract(Row row){
        this.row = row;
    }

    @Override
    public CellExtension createCell(int i) {
        return CellExtension.extendsCell(row.createCell(i));
    }

    @Override
    public CellExtension createCell(int i, CellType cellType) {
        return CellExtension.extendsCell(row.createCell(i, cellType));
    }

    @Override
    public void removeCell(Cell cell) {
        row.removeCell(cell);
    }

    @Override
    public void setRowNum(int i) {
        row.setRowNum(i);
    }

    @Override
    public int getRowNum() {
        return row.getRowNum();
    }

    @Override
    public CellExtension getCell(int i) {
        return row.getCell(i) == null ? null : CellExtension.extendsCell(row.getCell(i));
    }

    @Override
    public CellExtension getCell(int i, MissingCellPolicy missingCellPolicy) {
        return row.getCell(i, missingCellPolicy) == null ? null : CellExtension.extendsCell(row.getCell(i, missingCellPolicy));
    }

    @Override
    public short getFirstCellNum() {
        return row.getFirstCellNum();
    }

    @Override
    public short getLastCellNum() {
        return row.getLastCellNum();
    }

    @Override
    public int getPhysicalNumberOfCells() {
        return row.getPhysicalNumberOfCells();
    }

    @Override
    public void setHeight(short i) {
        row.setHeight(i);
    }

    @Override
    public void setZeroHeight(boolean b) {
        row.setZeroHeight(b);
    }

    @Override
    public boolean getZeroHeight() {
        return row.getZeroHeight();
    }

    @Override
    public void setHeightInPoints(float v) {
        row.setHeightInPoints(v);
    }

    @Override
    public short getHeight() {
        return row.getHeight();
    }

    @Override
    public float getHeightInPoints() {
        return row.getHeightInPoints();
    }

    @Override
    public boolean isFormatted() {
        return row.isFormatted();
    }

    @Override
    public CellStyle getRowStyle() {
        return row.getRowStyle();
    }

    @Override
    public void setRowStyle(CellStyle cellStyle) {
        row.setRowStyle(cellStyle);
    }

    @Override
    public Iterator<Cell> cellIterator() {
        return row.cellIterator();
    }

    @Override
    public SheetExtension getSheet() {
        return SheetExtension.extendsSheet(row.getSheet());
    }

    @Override
    public int getOutlineLevel() {
        return row.getOutlineLevel();
    }

    @Override
    public void shiftCellsRight(int i, int i1, int i2) {
        row.shiftCellsRight(i, i1, i2);
    }

    @Override
    public void shiftCellsLeft(int i, int i1, int i2) {
        row.shiftCellsLeft(i, i1, i2);
    }

    @Override
    public Iterator<Cell> iterator() {
        return row.iterator();
    }
}
