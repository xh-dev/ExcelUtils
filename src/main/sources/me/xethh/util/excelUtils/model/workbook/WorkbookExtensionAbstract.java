package me.xethh.util.excelUtils.model.workbook;

import me.xethh.util.excelUtils.model.sheet.SheetExtension;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.EvaluationWorkbook;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

public abstract class WorkbookExtensionAbstract implements WorkbookExtension{
    protected final Workbook workbook;
    protected WorkbookExtensionAbstract(Workbook workbook){
        this.workbook = workbook;
    }
    @Override
    public int getActiveSheetIndex() {
        return workbook.getActiveSheetIndex();
    }

    @Override
    public void setActiveSheet(int i) {
        workbook.setActiveSheet(i);
    }

    @Override
    public int getFirstVisibleTab() {
        return workbook.getFirstVisibleTab();
    }

    @Override
    public void setFirstVisibleTab(int i) {
        workbook.setFirstVisibleTab(i);
    }

    @Override
    public void setSheetOrder(String s, int i) {
        workbook.setSheetOrder(s, i);
    }

    @Override
    public void setSelectedTab(int i) {
        workbook.setSelectedTab(i);
    }

    @Override
    public void setSheetName(int i, String s) {
        workbook.setSheetName(i, s);
    }

    @Override
    public String getSheetName(int i) {
        return workbook.getSheetName(i);
    }

    @Override
    public int getSheetIndex(String s) {
        return workbook.getSheetIndex(s);
    }

    @Override
    public int getSheetIndex(Sheet sheet) {
        return workbook.getSheetIndex(sheet);
    }

    @Override
    public SheetExtension createSheet() {
        return SheetExtension.extendsSheet(workbook.createSheet());
    }

    @Override
    public SheetExtension createSheet(String s) {
        return SheetExtension.extendsSheet(workbook.createSheet(s));
    }

    @Override
    public SheetExtension cloneSheet(int i) {
        return SheetExtension.extendsSheet(workbook.cloneSheet(i));
    }

    @Override
    public Iterator<Sheet> sheetIterator() {
        return workbook.sheetIterator();
    }

    @Override
    public int getNumberOfSheets() {
        return workbook.getNumberOfSheets();
    }

    @Override
    public SheetExtension getSheetAt(int i) {
        return SheetExtension.extendsSheet(workbook.getSheetAt(i));
    }

    @Override
    public SheetExtension getSheet(String s) {
        return SheetExtension.extendsSheet(workbook.getSheet(s));
    }

    @Override
    public void removeSheetAt(int i) {
        workbook.removeSheetAt(i);
    }

    @Override
    public Font createFont() {
        return workbook.createFont();
    }

    @Override
    public Font findFont(boolean b, short i, short i1, String s, boolean b1, boolean b2, short i2, byte b3) {
        return workbook.findFont(b,i,i1,s,b1,b2,i2,b3);
    }

    @Deprecated
    @Override
    public int getNumberOfFonts() {
        return workbook.getNumberOfFonts();
    }

    @Override
    public int getNumberOfFontsAsInt() {
        return workbook.getNumberOfFontsAsInt();
    }

    @Override
    public Font getFontAt(int i) {
        return workbook.getFontAt(i);
    }

    @Override
    public CellStyle createCellStyle() {
        return workbook.createCellStyle();
    }

    @Override
    public int getNumCellStyles() {
        return workbook.getNumCellStyles();
    }

    @Override
    public CellStyle getCellStyleAt(int i) {
        return workbook.getCellStyleAt(i);
    }

    @Override
    public void write(OutputStream outputStream) throws IOException {
        workbook.write(outputStream);
    }

    @Override
    public void close() throws IOException {
        workbook.close();
    }

    @Override
    public int getNumberOfNames() {
        return workbook.getNumberOfNames();
    }

    @Override
    public Name getName(String s) {
        return workbook.getName(s);
    }

    @Override
    public List<? extends Name> getNames(String s) {
        return workbook.getNames(s);
    }

    @Override
    public List<? extends Name> getAllNames() {
        return workbook.getAllNames();
    }

    @Override
    public Name createName() {
        return workbook.createName();
    }

    @Override
    public void removeName(Name name) {
        workbook.removeName(name);
    }

    @Override
    public int linkExternalWorkbook(String s, Workbook workbook) {
        return workbook.linkExternalWorkbook(s, workbook);
    }

    @Override
    public void setPrintArea(int i, String s) {
        workbook.setPrintArea(i, s);
    }

    @Override
    public void setPrintArea(int i, int i1, int i2, int i3, int i4) {
        workbook.setPrintArea(i, i1, i2, i3, i4);
    }

    @Override
    public String getPrintArea(int i) {
        return workbook.getPrintArea(i);
    }

    @Override
    public void removePrintArea(int i) {
        workbook.removePrintArea(i);
    }

    @Override
    public Row.MissingCellPolicy getMissingCellPolicy() {
        return workbook.getMissingCellPolicy();
    }

    @Override
    public void setMissingCellPolicy(Row.MissingCellPolicy missingCellPolicy) {
        workbook.setMissingCellPolicy(missingCellPolicy);
    }

    @Override
    public DataFormat createDataFormat() {
        return workbook.createDataFormat();
    }

    @Override
    public int addPicture(byte[] bytes, int i) {
        return workbook.addPicture(bytes, i);
    }

    @Override
    public List<? extends PictureData> getAllPictures() {
        return workbook.getAllPictures();
    }

    @Override
    public CreationHelper getCreationHelper() {
        return workbook.getCreationHelper();
    }

    @Override
    public boolean isHidden() {
        return workbook.isHidden();
    }

    @Override
    public void setHidden(boolean b) {
        workbook.setHidden(b);
    }

    @Override
    public boolean isSheetHidden(int i) {
        return workbook.isSheetHidden(i);
    }

    @Override
    public boolean isSheetVeryHidden(int i) {
        return workbook.isSheetVeryHidden(i);
    }

    @Override
    public void setSheetHidden(int i, boolean b) {
        workbook.setSheetHidden(i, b);
    }

    @Override
    public SheetVisibility getSheetVisibility(int i) {
        return workbook.getSheetVisibility(i);
    }

    @Override
    public void setSheetVisibility(int i, SheetVisibility sheetVisibility) {
        workbook.setSheetVisibility(i, sheetVisibility);
    }

    @Override
    public void addToolPack(UDFFinder udfFinder) {
        workbook.addToolPack(udfFinder);
    }

    @Override
    public void setForceFormulaRecalculation(boolean b) {
        workbook.setForceFormulaRecalculation(b);
    }

    @Override
    public boolean getForceFormulaRecalculation() {
        return workbook.getForceFormulaRecalculation();
    }

    @Override
    public SpreadsheetVersion getSpreadsheetVersion() {
        return workbook.getSpreadsheetVersion();
    }

    @Override
    public int addOlePackage(byte[] bytes, String s, String s1, String s2) throws IOException {
        return workbook.addOlePackage(bytes, s, s1, s2);
    }

    @Override
    public EvaluationWorkbook createEvaluationWorkbook() {
        return workbook.createEvaluationWorkbook();
    }

    @Override
    public Iterator<Sheet> iterator() {
        return workbook.iterator();
    }
}
