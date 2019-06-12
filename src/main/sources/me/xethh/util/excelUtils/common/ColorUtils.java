package me.xethh.util.excelUtils.common;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.math.BigDecimal;

public class ColorUtils {
    enum TempCTColorType{Empty, IndexedColor, XSSFColor, HSSFColor}
    public static class TempCTColor{
        private TempCTColorType type;
        private int a,r,g,b;
        private BigDecimal tint;

        @Override
        public String toString() {
            return "TempCTColor{" +
                    "a=" + a +
                    ", r=" + r +
                    ", g=" + g +
                    ", b=" + b +
                    ", tint=" + tint +
                    '}';
        }

        public TempCTColorType getType() {
            return type;
        }

        public void setType(TempCTColorType type) {
            this.type = type;
        }

        public int getA() {
            return a;
        }

        public void setA(int a) {
            this.a = a;
        }

        public int getR() {
            return r;
        }

        public void setR(int r) {
            this.r = r;
        }

        public int getG() {
            return g;
        }

        public void setG(int g) {
            this.g = g;
        }

        public int getB() {
            return b;
        }

        public void setB(int b) {
            this.b = b;
        }

        public BigDecimal getTint() {
            return tint;
        }

        public void setTint(BigDecimal tint) {
            this.tint = tint;
        }
    }

    public static int applyTint(int lum, double tint){
        if(tint > 0){
            return (int) (lum * (1.0-tint) + (255 - 255 * (1.0-tint)));
        } else if (tint < 0){
            return (int) (lum*(1+tint));
        } else {
            return lum;
        }
    }
    public static TempCTColor toTempColor(Workbook workbook, short code, Color color){
        if(color instanceof XSSFColor && workbook instanceof XSSFWorkbook){
            TempCTColor ctColor = new TempCTColor();
            if(((XSSFColor) color).isIndexed()){
                IndexedColors c = IndexedColors.fromInt(code);
                if(IndexedColors.AUTOMATIC==c){
                    ctColor.setType(TempCTColorType.Empty);
                    return ctColor;
                }
                // color = new XSSFColor(c, ((XSSFWorkbook) workbook).getStylesSource().getIndexedColors());
            }
            byte[] argb = ((XSSFColor) color).getARGB();
            ctColor.a = argb[0] & 0xFF;
            ctColor.r = argb[1] & 0xFF;
            ctColor.g = argb[2] & 0xFF;
            ctColor.b = argb[3] & 0xFF;
            ctColor.tint = ((XSSFColor) color).hasTint()?new BigDecimal(((XSSFColor) color).getTint()+""):BigDecimal.ZERO;
            return ctColor;
        }
        if(color instanceof HSSFColor){
            TempCTColor ctColor = new TempCTColor();
            short[] argb = ((HSSFColor) color).getTriplet();
            ctColor.r = argb[0];
            ctColor.g = argb[1];
            ctColor.g = argb[2];
            ctColor.tint = ((XSSFColor) color).hasTint()?new BigDecimal(((XSSFColor) color).getTint()+""):BigDecimal.ZERO;
            return ctColor;
        }
        throw new RuntimeException("Unsupported type: "+color.getClass());
    }
}
