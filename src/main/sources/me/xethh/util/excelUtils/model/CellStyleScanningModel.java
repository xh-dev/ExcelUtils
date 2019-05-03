package me.xethh.util.excelUtils.model;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class CellStyleScanningModel {
    private String dataFromat;
    private Font font;
    private boolean isHiden;

    public String getDataFromat() {
        return dataFromat;
    }

    public void setDataFromat(String dataFromat) {
        this.dataFromat = dataFromat;
    }

    public Font getFont() {
        return font;
    }

    public void setFont(Font font) {
        this.font = font;
    }

    public boolean isHiden() {
        return isHiden;
    }

    public void setHiden(boolean hiden) {
        isHiden = hiden;
    }

    public CellStyleScanningModel(Workbook workbook, CellStyle cellStyle) {
        this.dataFromat = cellStyle.getDataFormatString();
        this.font = new Font(workbook.getFontAt(cellStyle.getFontIndexAsInt()));
        this.isHiden = cellStyle.getHidden();
    }

    public class Font{
        public Font(org.apache.poi.ss.usermodel.Font font) {
            this.name = font.getFontName();
            this.fontHeightInPoint = font.getFontHeightInPoints();
            this.isItalic = font.getItalic();
            this.isStrikeout = font.getStrikeout();
            this.color = font.getColor();
            this.typeOffset = font.getTypeOffset();
            this.isUnderLine = font.getUnderline();
            this.charSet = font.getCharSet();
            this.isBold = font.getBold();

        }

        private String name;
        private short fontHeightInPoint;
        private boolean isItalic;
        private boolean isStrikeout;
        private short color;
        private short typeOffset;
        private byte isUnderLine;
        private int charSet;
        private boolean isBold;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public short getFontHeightInPoint() {
            return fontHeightInPoint;
        }

        public void setFontHeightInPoint(short fontHeightInPoint) {
            this.fontHeightInPoint = fontHeightInPoint;
        }

        public boolean isItalic() {
            return isItalic;
        }

        public void setItalic(boolean italic) {
            isItalic = italic;
        }

        public boolean isStrikeout() {
            return isStrikeout;
        }

        public void setStrikeout(boolean strikeout) {
            isStrikeout = strikeout;
        }

        public short getColor() {
            return color;
        }

        public void setColor(short color) {
            this.color = color;
        }

        public short getTypeOffset() {
            return typeOffset;
        }

        public void setTypeOffset(short typeOffset) {
            this.typeOffset = typeOffset;
        }

        public byte isUnderLine() {
            return isUnderLine;
        }

        public void setUnderLine(byte underLine) {
            isUnderLine = underLine;
        }

        public int getCharSet() {
            return charSet;
        }

        public void setCharSet(int charSet) {
            this.charSet = charSet;
        }

        public boolean isBold() {
            return isBold;
        }

        public void setBold(boolean bold) {
            isBold = bold;
        }

        @Override
        public String toString() {
            return "Font{" +
                    "name='" + name + '\'' +
                    ", fontHeightInPoint=" + fontHeightInPoint +
                    ", isItalic=" + isItalic +
                    ", isStrikeout=" + isStrikeout +
                    ", color=" + color +
                    ", typeOffset=" + typeOffset +
                    ", isUnderLine=" + isUnderLine +
                    ", charSet=" + charSet +
                    ", isBold=" + isBold +
                    '}';
        }
    }

    @Override
    public String toString() {
        return "CellStyleScanningModel{" +
                "dataFromat='" + dataFromat + '\'' +
                ", font=" + font +
                '}';
    }
}
