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
    public static void main(String[] args) throws IOException {
        InputStream is = new FileInputStream(new File("./src/test/resources/TestingBase.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        // Iterator<CellScanningModel> scanning = WorkbookScanning.scan(workbook);
        // while(scanning.hasNext()){
        //     System.out.println(scanning.next());
        // }
        Iterator<String[]> scanning = WorkbookScanning.scanAsArr(workbook);
        // while(scanning.hasNext()){
        //     for(String s:scanning.next())
        //         System.out.print(s+", ");
        //     System.out.println();
        // }
        // FileOutputStream os = new FileOutputStream(new File("./src/test/resources/TestingBase_revised.xlsx"));
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("abc");
        int row = 0;
        while(scanning.hasNext()){
            String[] rowRecord = scanning.next();
            XSSFRow sheetRow = sheet.createRow(row);
            for(int i=0;i<rowRecord.length;i++) {
                sheetRow.createCell(i).setCellValue(rowRecord[i]);
            }
            if(rowRecord[27]!=null && !rowRecord[27].equals("") && rowRecord[27].split(",").length==3){
                int[] fontColor = toIntArr(rowRecord[27].split(","));
                int sum = 0;
                for(int s : fontColor)
                    sum+=s;
                if(sum>0){
                    XSSFCellStyle style = sheetRow.getCell(27).getCellStyle();
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(fontColor[0],fontColor[1],fontColor[2]),wb.getStylesSource().getIndexedColors()));
                    sheetRow.getCell(27).setCellStyle(style);
                }
            }
            if(rowRecord[8]!=null && !rowRecord[8].equals("") && rowRecord[8].split(",").length==4){
                int[] fontColor = toIntArr(rowRecord[8].split(","));
                int sum = 0;
                for(int s : fontColor)
                    sum+=s;
                if(sum>0){
                    XSSFCellStyle style = sheetRow.getCell(8).getCellStyle();
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(fontColor[0],fontColor[1],fontColor[2],fontColor[3]),wb.getStylesSource().getIndexedColors()));
                    sheetRow.getCell(8).setCellStyle(style);
                }
            }
            if(rowRecord[19]!=null && !rowRecord[19].equals("") && rowRecord[19].split(",").length==4){
                int[] fontColor = toIntArr(rowRecord[19].split(","));
                int sum = 0;
                for(int s : fontColor)
                    sum+=s;
                if(sum>0){
                    XSSFCellStyle style = sheetRow.getCell(19).getCellStyle();
                    style.setFillForegroundColor(new XSSFColor(new java.awt.Color(fontColor[0],fontColor[1],fontColor[2],fontColor[3]),wb.getStylesSource().getIndexedColors()));
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    sheetRow.getCell(19).setCellStyle(style);
                }
            }
            row++;
        }
        try {
            FileOutputStream outputStream = new FileOutputStream("./target/TestingBase_revised.xlsx");
            wb.write(outputStream);
            wb.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
