package me.xethh.util.excelUtils.model.col;

import me.xethh.util.excelUtils.model.row.RowExtension;
import me.xethh.util.excelUtils.model.sheet.SheetExtension;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

public class CellExtensionAbstract implements CellExtension{
    protected Cell cell;
    protected CellExtensionAbstract(Cell cell){
        this.cell = cell;
    }
    @Override
    public int getColumnIndex() {
        return cell.getColumnIndex();
    }

    @Override
    public int getRowIndex() {
        return cell.getRowIndex();
    }

    @Override
    public SheetExtension getSheet() {
        return SheetExtension.extendsSheet(cell.getSheet());
    }

    @Override
    public RowExtension getRow() {
        return RowExtension.extendsRow(cell.getRow());
    }

    @Deprecated
    @Override
    public void setCellType(CellType cellType) {
        cell.setCellType(cellType);
    }

    @Override
    public void setBlank() {
        cell.setBlank();
    }

    @Override
    public CellType getCellType() {
        return cell.getCellType();
    }

    @Override
    public CellType getCachedFormulaResultType() {
        return cell.getCachedFormulaResultType();
    }

    @Override
    public void setCellValue(double v) {
        cell.setCellValue(v);
    }

    @Override
    public void setCellValue(Date date) {
        cell.setCellValue(date);
    }

    @Override
    public void setCellValue(LocalDateTime value) {
        cell.setCellValue(value);
    }

    @Override
    public void setCellValue(Calendar calendar) {
        cell.setCellValue(calendar);
    }

    @Override
    public void setCellValue(RichTextString richTextString) {
        cell.setCellValue(richTextString);
    }

    @Override
    public void setCellValue(String s) {
        cell.setCellValue(s);
    }

    @Override
    public void setCellFormula(String s) throws FormulaParseException, IllegalStateException {
        cell.setCellValue(s);
    }

    @Override
    public void removeFormula() throws IllegalStateException {
        cell.removeFormula();
    }

    @Override
    public String getCellFormula() {
        return cell.getCellFormula();
    }

    @Override
    public double getNumericCellValue() {
        return cell.getNumericCellValue();
    }

    @Override
    public Date getDateCellValue() {
        return cell.getDateCellValue();
    }

    @Override
    public LocalDateTime getLocalDateTimeCellValue() {
        return cell.getLocalDateTimeCellValue();
    }

    @Override
    public RichTextString getRichStringCellValue() {
        return cell.getRichStringCellValue();
    }

    @Override
    public String getStringCellValue() {
        return cell.getStringCellValue();
    }

    @Override
    public void setCellValue(boolean b) {
        cell.setCellValue(b);
    }

    @Override
    public void setCellErrorValue(byte b) {
        cell.setCellErrorValue(b);
    }

    @Override
    public boolean getBooleanCellValue() {
        return cell.getBooleanCellValue();
    }

    @Override
    public byte getErrorCellValue() {
        return cell.getErrorCellValue();
    }

    @Override
    public void setCellStyle(CellStyle cellStyle) {
        cell.setCellStyle(cellStyle);
    }

    @Override
    public CellStyle getCellStyle() {
        return cell.getCellStyle();
    }

    @Override
    public void setAsActiveCell() {
        cell.setAsActiveCell();
    }

    @Override
    public CellAddress getAddress() {
        return cell.getAddress();
    }

    @Override
    public void setCellComment(Comment comment) {
        cell.setCellComment(comment);
    }

    @Override
    public Comment getCellComment() {
        return cell.getCellComment();
    }

    @Override
    public void removeCellComment() {
        cell.removeCellComment();
    }

    @Override
    public Hyperlink getHyperlink() {
        return cell.getHyperlink();
    }

    @Override
    public void setHyperlink(Hyperlink hyperlink) {
        cell.setHyperlink(hyperlink);
    }

    @Override
    public void removeHyperlink() {
        cell.removeHyperlink();
    }

    @Override
    public CellRangeAddress getArrayFormulaRange() {
        return cell.getArrayFormulaRange();
    }

    @Override
    public boolean isPartOfArrayFormulaGroup() {
        return cell.isPartOfArrayFormulaGroup();
    }
}
