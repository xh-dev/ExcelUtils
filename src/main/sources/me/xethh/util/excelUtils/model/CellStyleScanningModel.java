package me.xethh.util.excelUtils.model;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.Arrays;

public class CellStyleScanningModel {
    private String dataFromat;
    private Font font;
    private boolean isHidden;
    private boolean isLocked;
    private boolean isQuotedPrefix;
    private HorizontalAlignment horizontalAlignment;
    private boolean wrappedText;
    private VerticalAlignment verticalAlignment;
    private short rotation;
    private short indentation;
    private BorderStyle borderLeft, borderRight, borderTop, borderBot;
    private IndexedColors borderLeftColor, borderRightColor, borderTopColor, borderBotColor;
    private FillPatternType fillPatternType;
    private IndexedColors backgroundColor;
    private short[] backgroundRGB;
    private IndexedColors foregroundColor;
    private short[] foregroundRGB;
    private boolean shrinkToFit;


    public CellStyleScanningModel(Workbook workbook, CellStyle cellStyle) {
        this.dataFromat = cellStyle.getDataFormatString();
        this.font = new Font(workbook, workbook.getFontAt(cellStyle.getFontIndexAsInt()));
        this.isHidden = cellStyle.getHidden();
        this.isLocked = cellStyle.getLocked();
        this.isQuotedPrefix = cellStyle.getQuotePrefixed();
        this.horizontalAlignment = cellStyle.getAlignment();
        this.wrappedText = cellStyle.getWrapText();
        this.verticalAlignment = cellStyle.getVerticalAlignment();
        this.rotation = cellStyle.getRotation();
        this.indentation= cellStyle.getIndention();
        this.borderLeft = cellStyle.getBorderLeft();
        this.borderRight = cellStyle.getBorderRight();
        this.borderTop = cellStyle.getBorderTop();
        this.borderBot = cellStyle.getBorderBottom();
        this.borderLeftColor = IndexedColors.fromInt(cellStyle.getLeftBorderColor());
        this.borderRightColor = IndexedColors.fromInt(cellStyle.getRightBorderColor());
        this.borderTopColor = IndexedColors.fromInt(cellStyle.getTopBorderColor());
        this.borderBotColor = IndexedColors.fromInt(cellStyle.getBottomBorderColor());
        this.fillPatternType = cellStyle.getFillPattern();
        this.backgroundColor = IndexedColors.fromInt(cellStyle.getFillBackgroundColor());
        this.foregroundColor = IndexedColors.fromInt(cellStyle.getFillForegroundColor());

        if(cellStyle.getFillBackgroundColorColor()!=null && cellStyle.getFillBackgroundColorColor() instanceof XSSFColor){
            if(((XSSFColor) cellStyle.getFillBackgroundColorColor()).getRGB()!=null){
                byte[] rgb = ((XSSFColor) cellStyle.getFillBackgroundColorColor()).getRGB();
                this.foregroundRGB = new short[rgb.length];
                for(int i=0;i<this.foregroundRGB.length;i++)
                    this.foregroundRGB[i] = (short) (rgb[i] & 0xff);
            }
        }
        else if(cellStyle.getFillBackgroundColorColor()!=null && cellStyle.getFillBackgroundColorColor() instanceof HSSFColor){
            if(((HSSFColor) cellStyle.getFillBackgroundColorColor()).getTriplet()!=null){
                this.foregroundRGB = ((HSSFColor) cellStyle.getFillBackgroundColorColor()).getTriplet();
            }
        }
        if(cellStyle.getFillBackgroundColorColor()!=null && cellStyle.getFillBackgroundColorColor() instanceof XSSFColor){
            if(((XSSFColor) cellStyle.getFillForegroundColorColor()).getRGB()!=null){
                byte[] rgb = ((XSSFColor) cellStyle.getFillForegroundColorColor()).getRGB();
                this.foregroundRGB = new short[rgb.length];
                for(int i=0;i<this.foregroundRGB.length;i++)
                    this.foregroundRGB[i] = (short) (rgb[i] & 0xff);
            }

        }
        else if(cellStyle.getFillBackgroundColorColor()!=null && cellStyle.getFillBackgroundColorColor() instanceof HSSFColor){
            if(((HSSFColor) cellStyle.getFillForegroundColorColor()).getTriplet()!=null){
                this.foregroundRGB = ((HSSFColor) cellStyle.getFillBackgroundColorColor()).getTriplet();
            }
        }
        Color fo = cellStyle.getFillForegroundColorColor();

        this.shrinkToFit = cellStyle.getShrinkToFit();
    }

    public class Font{
        public Font(Workbook workbook, org.apache.poi.ss.usermodel.Font font) {
            this.name = font.getFontName();
            this.fontHeightInPoint = font.getFontHeightInPoints();
            this.isItalic = font.getItalic();
            this.isStrikeout = font.getStrikeout();
            this.color = font.getColor();

            if (null != font) {
                if (font instanceof XSSFFont) {
                    XSSFColor temp = ((XSSFFont) font).getXSSFColor();
                    if (null != temp) {
                        byte[] rgb = temp.getRGB();
                        this.colorRgb = new short[temp.getRGB().length];
                        for (int i = 0; i < rgb.length; i++) {
                            this.colorRgb[i] = (short) (rgb[i] & 0xff);
                        }
                    }
                }
                if(font instanceof HSSFFont){
                    HSSFColor tempColor = ((HSSFFont) font).getHSSFColor((HSSFWorkbook) workbook);
                    if(null != tempColor){
                        this.colorRgb = tempColor.getTriplet();
                    }

                }
            }

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
        private short[] colorRgb;
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

        public short[] getColorRgb() {
            return colorRgb;
        }

        public void setColorRgb(short[] colorRgb) {
            this.colorRgb = colorRgb;
        }

        public byte getIsUnderLine() {
            return isUnderLine;
        }

        public void setIsUnderLine(byte isUnderLine) {
            this.isUnderLine = isUnderLine;
        }

        @Override
        public String toString() {
            return "Font{" +
                    "name='" + name + '\'' +
                    ", fontHeightInPoint=" + fontHeightInPoint +
                    ", isItalic=" + isItalic +
                    ", isStrikeout=" + isStrikeout +
                    ", color=" + color +
                    ", colorRgb=" + Arrays.toString(colorRgb) +
                    ", typeOffset=" + typeOffset +
                    ", isUnderLine=" + isUnderLine +
                    ", charSet=" + charSet +
                    ", isBold=" + isBold +
                    '}';
        }
    }

    public boolean isLocked() {
        return isLocked;
    }

    public void setLocked(boolean locked) {
        isLocked = locked;
    }

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

    public boolean isHidden() {
        return isHidden;
    }

    public void setHidden(boolean hidden) {
        isHidden = hidden;
    }

    public boolean isQuotedPrefix() {
        return isQuotedPrefix;
    }

    public void setQuotedPrefix(boolean quotedPrefix) {
        isQuotedPrefix = quotedPrefix;
    }

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public void setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public boolean isWrappedText() {
        return wrappedText;
    }

    public void setWrappedText(boolean wrappedText) {
        this.wrappedText = wrappedText;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public short getRotation() {
        return rotation;
    }

    public void setRotation(short rotation) {
        this.rotation = rotation;
    }

    public short getIndentation() {
        return indentation;
    }

    public void setIndentation(short indentation) {
        this.indentation = indentation;
    }

    public BorderStyle getBorderLeft() {
        return borderLeft;
    }

    public void setBorderLeft(BorderStyle borderLeft) {
        this.borderLeft = borderLeft;
    }

    public BorderStyle getBorderRight() {
        return borderRight;
    }

    public void setBorderRight(BorderStyle borderRight) {
        this.borderRight = borderRight;
    }

    public BorderStyle getBorderTop() {
        return borderTop;
    }

    public void setBorderTop(BorderStyle borderTop) {
        this.borderTop = borderTop;
    }

    public BorderStyle getBorderBot() {
        return borderBot;
    }

    public void setBorderBot(BorderStyle borderBot) {
        this.borderBot = borderBot;
    }

    public IndexedColors getBorderLeftColor() {
        return borderLeftColor;
    }

    public void setBorderLeftColor(IndexedColors borderLeftColor) {
        this.borderLeftColor = borderLeftColor;
    }

    public IndexedColors getBorderRightColor() {
        return borderRightColor;
    }

    public void setBorderRightColor(IndexedColors borderRightColor) {
        this.borderRightColor = borderRightColor;
    }

    public IndexedColors getBorderTopColor() {
        return borderTopColor;
    }

    public void setBorderTopColor(IndexedColors borderTopColor) {
        this.borderTopColor = borderTopColor;
    }

    public IndexedColors getBorderBotColor() {
        return borderBotColor;
    }

    public void setBorderBotColor(IndexedColors borderBotColor) {
        this.borderBotColor = borderBotColor;
    }

    public FillPatternType getFillPatternType() {
        return fillPatternType;
    }

    public void setFillPatternType(FillPatternType fillPatternType) {
        this.fillPatternType = fillPatternType;
    }

    public IndexedColors getBackgroundColor() {
        return backgroundColor;
    }

    public void setBackgroundColor(IndexedColors backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    public IndexedColors getForegroundColor() {
        return foregroundColor;
    }

    public void setForegroundColor(IndexedColors foregroundColor) {
        this.foregroundColor = foregroundColor;
    }

    public boolean isShrinkToFit() {
        return shrinkToFit;
    }

    public void setShrinkToFit(boolean shrinkToFit) {
        this.shrinkToFit = shrinkToFit;
    }

    public short[] getBackgroundRGB() {
        return backgroundRGB;
    }

    public void setBackgroundRGB(short[] backgroundRGB) {
        this.backgroundRGB = backgroundRGB;
    }

    public short[] getForegroundRGB() {
        return foregroundRGB;
    }

    public void setForegroundRGB(short[] foregroundRGB) {
        this.foregroundRGB = foregroundRGB;
    }

    @Override
    public String toString() {
        return "CellStyleScanningModel{" +
                "dataFromat='" + dataFromat + '\'' +
                ", font=" + font +
                ", isHidden=" + isHidden +
                ", isLocked=" + isLocked +
                ", isQuotedPrefix=" + isQuotedPrefix +
                ", horizontalAlignment=" + horizontalAlignment +
                ", wrappedText=" + wrappedText +
                ", verticalAlignment=" + verticalAlignment +
                ", rotation=" + rotation +
                ", indentation=" + indentation +
                ", borderLeft=" + borderLeft +
                ", borderRight=" + borderRight +
                ", borderTop=" + borderTop +
                ", borderBot=" + borderBot +
                ", borderLeftColor=" + borderLeftColor +
                ", borderRightColor=" + borderRightColor +
                ", borderTopColor=" + borderTopColor +
                ", borderBotColor=" + borderBotColor +
                ", fillPatternType=" + fillPatternType +
                ", backgroundColor=" + backgroundColor +
                ", backgroundRGB=" + Arrays.toString(backgroundRGB) +
                ", foregroundColor=" + foregroundColor +
                ", foregroundRGB=" + Arrays.toString(foregroundRGB) +
                ", shrinkToFit=" + shrinkToFit +
                '}';
    }
}
