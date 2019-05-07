package me.xethh.util.excelUtils;

import static org.junit.Assert.assertTrue;

import me.xethh.util.excelUtils.common.ExcelReadValue;
import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.util.excelUtils.model.CellStyleScanningModel;
import me.xethh.util.excelUtils.reading.WorkbookScanning;
import me.xethh.utils.wrapper.Tuple2;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.io.*;
import java.util.Iterator;

/**
 * Unit test for simple App.
 */
public class AppTest 
{
    /**
     * Rigorous Test :-)
     */
    @Test
    public void shouldAnswerWithTrue()
    {
        assertTrue( true );
    }

    public static int[] toIntArr(String[] ss){
        int[] is = new int[ss.length];
        for(int i=0;i<ss.length;i++){
            is[i] = Integer.parseInt(ss[i]);
        }
        return is;
    }
    public static byte[] toByteArr(String[] ss){
        int[] is = toIntArr(ss);
        byte[] ba = new byte[ss.length];
        for(int i=0;i<is.length;i++){
            ba[i] = (byte) is[i];
        }
        return ba;
    }
    public static void main(String[] args) throws IOException {
        InputStream is = new FileInputStream(new File("./src/test/resources/TestingBase.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        Iterator<String[]> scanning = WorkbookScanning.scanAsArr(workbook);
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("abc");
        int row = 0;
        while(scanning.hasNext()){
            String[] rowRecord = scanning.next();
            XSSFRow sheetRow = sheet.createRow(row);
            for(int i=0;i<rowRecord.length;i++) {
                XSSFCell cell = sheetRow.createCell(i);
                if(i==27 && rowRecord[i]!=null && !rowRecord[i].equals("") && rowRecord[i].split(",").length==3){
                    int[] fontColor = toIntArr(rowRecord[i].split(","));
                    byte[] fontColorByte = toByteArr(rowRecord[i].split(","));
                    int sum = 0;
                    for(int s : fontColor)
                        sum+=s;
                    if(sum>0){
                        XSSFCellStyle style = wb.createCellStyle();
                        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        style.setFillForegroundColor(new XSSFColor(fontColorByte,wb.getStylesSource().getIndexedColors()));
                        cell.setCellStyle(style);
                    }
                }
                if(i==8 && rowRecord[8]!=null && !rowRecord[8].equals("") && rowRecord[8].split(",").length==4){
                    int[] fontColor = toIntArr(rowRecord[8].split(","));
                    byte[] fontColorByte = toByteArr(rowRecord[i].split(","));
                    int sum = 0;
                    for(int s : fontColor)
                        sum+=s;
                    if(sum>0){
                        XSSFCellStyle style = wb.createCellStyle();
                        sheetRow.getCell(12).getCellStyle();
                        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        style.setFillForegroundColor(new XSSFColor(fontColorByte,wb.getStylesSource().getIndexedColors()));
                        sheetRow.getCell(8).setCellStyle(style);
                    }
                }
                if(i == 19 && rowRecord[19]!=null && !rowRecord[19].equals("") && rowRecord[19].split(",").length==4){
                    int[] fontColor = toIntArr(rowRecord[19].split(","));
                    byte[] fontColorByte = toByteArr(rowRecord[i].split(","));
                    int sum = 0;
                    for(int s : fontColor)
                        sum+=s;
                    if(sum>0){
                        XSSFCellStyle style = wb.createCellStyle();
                        style.setFillForegroundColor(new XSSFColor(fontColorByte,wb.getStylesSource().getIndexedColors()));
                        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        sheetRow.getCell(19).setCellStyle(style);
                    }
                }
                cell.setCellValue(rowRecord[i]);
            }
            row++;
        }
        try {
            String filePath = "./target/TestingBase_revised.xlsx";
            if(new File(filePath).exists()) new File(filePath).delete();
            FileOutputStream outputStream = new FileOutputStream(filePath);
            wb.write(outputStream);
            wb.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
