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
        setCellStr("");
    }

    public Integer getActCol() {
        return actCol;
    }

    public void setActCol(Integer actCol) {
        this.actCol = actCol;
        setCellStr("");
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

    /**
     * This is a dummy setter only
     * @param cellStr
     */
    public void setCellStr(String cellStr) {
        if(this.actRow!=null && this.actCol!=null && this.cellStr==null)
            this.cellStr = getCellStr();
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public static String[] toStringArrHeader(){
        String[] arr = new String[31];
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
    public String[] toStringArr(){
        String[] arr = new String[31];
        CellScanningModel cellModel = this;
        arr[0] = cellModel.getSheetName();
        arr[1] = String.valueOf(cellModel.getActRow());
        arr[2] = String.valueOf(cellModel.getActCol());
        arr[3] = cellModel.getCellStr();
        arr[4] = cellModel.getCellType().name();
        arr[5] = cellModel.getValue()==null?"[null]":cellModel.getValue().toString();
        CellStyleScanningModel style = cellModel.getCellStyle();
        arr[6] = style.getDataFromat();
        arr[7] = style.getBackgroundColor().name();

        arr[8] = style.getBackgroundRGB()==null?"":style.getBackgroundRGB().toString();
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
        arr[19] = style.getForegroundRGB()==null?"":style.getForegroundRGB().toString();
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

    private static String colorToString(short[] argb){
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
