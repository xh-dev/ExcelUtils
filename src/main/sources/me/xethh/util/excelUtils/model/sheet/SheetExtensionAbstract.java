package me.xethh.util.excelUtils.model.sheet;

import me.xethh.util.excelUtils.model.row.RowExtension;
import me.xethh.util.excelUtils.model.workbook.WorkbookExtension;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;

import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public abstract class SheetExtensionAbstract implements SheetExtension {
    protected final Sheet sheet;
    protected SheetExtensionAbstract(Sheet sheet){
        this.sheet = sheet;
    }
    @Override
    public RowExtension createRow(int i) {
        return RowExtension.extendsRow(sheet.createRow(i));
    }

    @Override
    public void removeRow(Row row) {
        sheet.removeRow(row);
    }

    @Override
    public RowExtension getRow(int i) {
        return sheet.getRow(i) == null? null :RowExtension.extendsRow(sheet.getRow(i));
    }

    @Override
    public int getPhysicalNumberOfRows() {
        return sheet.getPhysicalNumberOfRows();
    }

    @Override
    public int getFirstRowNum() {
        return sheet.getFirstRowNum();
    }

    @Override
    public int getLastRowNum() {
        return sheet.getLastRowNum();
    }

    @Override
    public void setColumnHidden(int i, boolean b) {
        sheet.setColumnHidden(i, b);
    }

    @Override
    public boolean isColumnHidden(int i) {
        return sheet.isColumnHidden(i);
    }

    @Override
    public void setRightToLeft(boolean b) {
        sheet.setRightToLeft(b);
    }

    @Override
    public boolean isRightToLeft() {
        return sheet.isRightToLeft();
    }

    @Override
    public void setColumnWidth(int i, int i1) {
        sheet.setColumnWidth(i, i1);
    }

    @Override
    public int getColumnWidth(int i) {
        return sheet.getColumnWidth(i);
    }

    @Override
    public float getColumnWidthInPixels(int i) {
        return sheet.getColumnWidthInPixels(i);
    }

    @Override
    public void setDefaultColumnWidth(int i) {
        sheet.setDefaultColumnWidth(i);
    }

    @Override
    public int getDefaultColumnWidth() {
        return sheet.getDefaultColumnWidth();
    }

    @Override
    public short getDefaultRowHeight() {
        return sheet.getDefaultRowHeight();
    }

    @Override
    public float getDefaultRowHeightInPoints() {
        return sheet.getDefaultRowHeightInPoints();
    }

    @Override
    public void setDefaultRowHeight(short i) {
        sheet.setDefaultRowHeight(i);
    }

    @Override
    public void setDefaultRowHeightInPoints(float v) {
        sheet.setDefaultRowHeightInPoints(v);
    }

    @Override
    public CellStyle getColumnStyle(int i) {
        return sheet.getColumnStyle(i);
    }

    @Override
    public int addMergedRegion(CellRangeAddress cellRangeAddress) {
        return sheet.addMergedRegion(cellRangeAddress);
    }

    @Override
    public int addMergedRegionUnsafe(CellRangeAddress cellRangeAddress) {
        return sheet.addMergedRegionUnsafe(cellRangeAddress);
    }

    @Override
    public void validateMergedRegions() {
        sheet.validateMergedRegions();
    }

    @Override
    public void setVerticallyCenter(boolean b) {
        sheet.setVerticallyCenter(b);
    }

    @Override
    public void setHorizontallyCenter(boolean b) {
        sheet.setHorizontallyCenter(b);
    }

    @Override
    public boolean getHorizontallyCenter() {
        return sheet.getHorizontallyCenter();
    }

    @Override
    public boolean getVerticallyCenter() {
        return sheet.getVerticallyCenter();
    }

    @Override
    public void removeMergedRegion(int i) {
        sheet.removeMergedRegion(i);
    }

    @Override
    public void removeMergedRegions(Collection<Integer> collection) {
        sheet.removeMergedRegions(collection);
    }

    @Override
    public int getNumMergedRegions() {
        return sheet.getNumMergedRegions();
    }

    @Override
    public CellRangeAddress getMergedRegion(int i) {
        return sheet.getMergedRegion(i);
    }

    @Override
    public List<CellRangeAddress> getMergedRegions() {
        return sheet.getMergedRegions();
    }

    @Override
    public Iterator<Row> rowIterator() {
        return sheet.iterator();
    }

    @Override
    public void setForceFormulaRecalculation(boolean b) {
        sheet.setForceFormulaRecalculation(b);
    }

    @Override
    public boolean getForceFormulaRecalculation() {
        return sheet.getForceFormulaRecalculation();
    }

    @Override
    public void setAutobreaks(boolean b) {
        sheet.setAutobreaks(b);
    }

    @Override
    public void setDisplayGuts(boolean b) {
        sheet.setDisplayGuts(b);
    }

    @Override
    public void setDisplayZeros(boolean b) {
        sheet.setDisplayZeros(b);
    }

    @Override
    public boolean isDisplayZeros() {
        return sheet.isDisplayZeros();
    }

    @Override
    public void setFitToPage(boolean b) {
        sheet.setFitToPage(b);
    }

    @Override
    public void setRowSumsBelow(boolean b) {
        sheet.setRowSumsBelow(b);
    }

    @Override
    public void setRowSumsRight(boolean b) {
        sheet.setRowSumsRight(b);
    }

    @Override
    public boolean getAutobreaks() {
        return sheet.getAutobreaks();
    }

    @Override
    public boolean getDisplayGuts() {
        return sheet.getDisplayGuts();
    }

    @Override
    public boolean getFitToPage() {
        return sheet.getFitToPage();
    }

    @Override
    public boolean getRowSumsBelow() {
        return sheet.getRowSumsBelow();
    }

    @Override
    public boolean getRowSumsRight() {
        return sheet.getRowSumsRight();
    }

    @Override
    public boolean isPrintGridlines() {
        return sheet.isPrintGridlines();
    }

    @Override
    public void setPrintGridlines(boolean b) {
        sheet.setPrintGridlines(b);
    }

    @Override
    public boolean isPrintRowAndColumnHeadings() {
        return sheet.isPrintRowAndColumnHeadings();
    }

    @Override
    public void setPrintRowAndColumnHeadings(boolean b) {
        sheet.setPrintRowAndColumnHeadings(b);
    }

    @Override
    public PrintSetup getPrintSetup() {
        return sheet.getPrintSetup();
    }

    @Override
    public Header getHeader() {
        return sheet.getHeader();
    }

    @Override
    public Footer getFooter() {
        return sheet.getFooter();
    }

    @Override
    public void setSelected(boolean b) {
        sheet.setSelected(b);
    }

    @Override
    public double getMargin(short i) {
        return sheet.getMargin(i);
    }

    @Override
    public void setMargin(short i, double v) {
        sheet.setMargin(i, v);
    }

    @Override
    public boolean getProtect() {
        return sheet.getProtect();
    }

    @Override
    public void protectSheet(String s) {
        sheet.protectSheet(s);
    }

    @Override
    public boolean getScenarioProtect() {
        return sheet.getScenarioProtect();
    }

    @Override
    public void setZoom(int i) {
        sheet.setZoom(i);
    }

    @Override
    public short getTopRow() {
        return sheet.getTopRow();
    }

    @Override
    public short getLeftCol() {
        return sheet.getLeftCol();
    }

    @Override
    public void showInPane(int i, int i1) {
        sheet.showInPane(i, i1);
    }

    @Override
    public void shiftRows(int i, int i1, int i2) {
        sheet.shiftRows(i, i1, i2);
    }

    @Override
    public void shiftRows(int i, int i1, int i2, boolean b, boolean b1) {
        sheet.shiftRows(i, i1, i2, b, b1);
    }

    @Override
    public void shiftColumns(int i, int i1, int i2) {
        sheet.shiftColumns(i, i1, i2);
    }

    @Override
    public void createFreezePane(int i, int i1, int i2, int i3) {
        sheet.createFreezePane(i, i1, i2, i3);
    }

    @Override
    public void createFreezePane(int i, int i1) {
        sheet.createFreezePane(i, i1);
    }

    @Override
    public void createSplitPane(int i, int i1, int i2, int i3, int i4) {
        sheet.createSplitPane(i, i1, i2, i3, i4);
    }

    @Override
    public PaneInformation getPaneInformation() {
        return sheet.getPaneInformation();
    }

    @Override
    public void setDisplayGridlines(boolean b) {
        sheet.setDisplayGridlines(b);
    }

    @Override
    public boolean isDisplayGridlines() {
        return sheet.isDisplayGridlines();
    }

    @Override
    public void setDisplayFormulas(boolean b) {
        sheet.setDisplayFormulas(b);
    }

    @Override
    public boolean isDisplayFormulas() {
        return sheet.isDisplayFormulas();
    }

    @Override
    public void setDisplayRowColHeadings(boolean b) {
        sheet.setDisplayRowColHeadings(b);
    }

    @Override
    public boolean isDisplayRowColHeadings() {
        return sheet.isDisplayRowColHeadings();
    }

    @Override
    public void setRowBreak(int i) {
        sheet.setRowBreak(i);
    }

    @Override
    public boolean isRowBroken(int i) {
        return sheet.isRowBroken(i);
    }

    @Override
    public void removeRowBreak(int i) {
        sheet.removeRowBreak(i);
    }

    @Override
    public int[] getRowBreaks() {
        return sheet.getRowBreaks();
    }

    @Override
    public int[] getColumnBreaks() {
        return sheet.getColumnBreaks();
    }

    @Override
    public void setColumnBreak(int i) {
        sheet.setColumnBreak(i);
    }

    @Override
    public boolean isColumnBroken(int i) {
        return sheet.isColumnBroken(i);
    }

    @Override
    public void removeColumnBreak(int i) {
        sheet.removeColumnBreak(i);
    }

    @Override
    public void setColumnGroupCollapsed(int i, boolean b) {
        sheet.setColumnGroupCollapsed(i, b);
    }

    @Override
    public void groupColumn(int i, int i1) {
        sheet.groupColumn(i, i1);
    }

    @Override
    public void ungroupColumn(int i, int i1) {
        sheet.ungroupColumn(i, i1);
    }

    @Override
    public void groupRow(int i, int i1) {
        sheet.groupRow(i, i1);
    }

    @Override
    public void ungroupRow(int i, int i1) {
        sheet.ungroupRow(i, i1);
    }

    @Override
    public void setRowGroupCollapsed(int i, boolean b) {
        sheet.setRowGroupCollapsed(i, b);
    }

    @Override
    public void setDefaultColumnStyle(int i, CellStyle cellStyle) {
        sheet.setDefaultColumnStyle(i, cellStyle);
    }

    @Override
    public void autoSizeColumn(int i) {
        sheet.autoSizeColumn(i);
    }

    @Override
    public void autoSizeColumn(int i, boolean b) {
        sheet.autoSizeColumn(i, b);
    }

    @Override
    public Comment getCellComment(CellAddress cellAddress) {
        return sheet.getCellComment(cellAddress);
    }

    @Override
    public Map<CellAddress, ? extends Comment> getCellComments() {
        return sheet.getCellComments();
    }

    @Override
    public Drawing<?> getDrawingPatriarch() {
        return sheet.getDrawingPatriarch();
    }

    @Override
    public Drawing<?> createDrawingPatriarch() {
        return sheet.createDrawingPatriarch();
    }

    @Override
    public WorkbookExtension getWorkbook() {
        return WorkbookExtension.extendsWorkbook(sheet.getWorkbook());
    }

    @Override
    public String getSheetName() {
        return sheet.getSheetName();
    }

    @Override
    public boolean isSelected() {
        return sheet.isSelected();
    }

    @Override
    public CellRange<? extends Cell> setArrayFormula(String s, CellRangeAddress cellRangeAddress) {
        return sheet.setArrayFormula(s, cellRangeAddress);
    }

    @Override
    public CellRange<? extends Cell> removeArrayFormula(Cell cell) {
        return sheet.removeArrayFormula(cell);
    }

    @Override
    public DataValidationHelper getDataValidationHelper() {
        return sheet.getDataValidationHelper();
    }

    @Override
    public List<? extends DataValidation> getDataValidations() {
        return sheet.getDataValidations();
    }

    @Override
    public void addValidationData(DataValidation dataValidation) {
        sheet.addValidationData(dataValidation);
    }

    @Override
    public AutoFilter setAutoFilter(CellRangeAddress cellRangeAddress) {
        return sheet.setAutoFilter(cellRangeAddress);
    }

    @Override
    public SheetConditionalFormatting getSheetConditionalFormatting() {
        return sheet.getSheetConditionalFormatting();
    }

    @Override
    public CellRangeAddress getRepeatingRows() {
        return sheet.getRepeatingRows();
    }

    @Override
    public CellRangeAddress getRepeatingColumns() {
        return sheet.getRepeatingColumns();
    }

    @Override
    public void setRepeatingRows(CellRangeAddress cellRangeAddress) {
        sheet.setRepeatingRows(cellRangeAddress);
    }

    @Override
    public void setRepeatingColumns(CellRangeAddress cellRangeAddress) {
        sheet.setRepeatingColumns(cellRangeAddress);
    }

    @Override
    public int getColumnOutlineLevel(int i) {
        return sheet.getColumnOutlineLevel(i);
    }

    @Override
    public Hyperlink getHyperlink(int i, int i1) {
        return sheet.getHyperlink(i, i1);
    }

    @Override
    public Hyperlink getHyperlink(CellAddress cellAddress) {
        return sheet.getHyperlink(cellAddress);
    }

    @Override
    public List<? extends Hyperlink> getHyperlinkList() {
        return sheet.getHyperlinkList();
    }

    @Override
    public CellAddress getActiveCell() {
        return sheet.getActiveCell();
    }

    @Override
    public void setActiveCell(CellAddress cellAddress) {
        sheet.setActiveCell(cellAddress);
    }

    @Override
    public Iterator<Row> iterator() {
        return sheet.iterator();
    }
}
