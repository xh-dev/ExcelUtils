package me.xethh.util.excelUtils.reading;

import me.xethh.util.excelUtils.common.ExcelReadValue;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.Date;
import java.util.Iterator;

public class ReadingFactory {
    private static void checkAndFillColor(XSSFWorkbook wb, Cell cell, int nowIndex, int index, String[] rowRecord){
        if(nowIndex==index && rowRecord[index]!=null && !rowRecord[index].equals("") && (rowRecord[index].split(",").length==3 || rowRecord[index].split(",").length==4)){
            byte[] fontColorByte = ExcelReadValue.toByteArr(rowRecord[index].split(","));
            if(!ExcelReadValue.isPureDark(rowRecord[index].split(","))){
                XSSFCellStyle style = wb.createCellStyle();
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setFillForegroundColor(new XSSFColor(fontColorByte,wb.getStylesSource().getIndexedColors()));
                cell.setCellStyle(style);
            }
        }
    }
    public static void scanExcel(String source, String dest) throws IOException {
        InputStream is = new FileInputStream(new File(source));
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        Iterator<String[]> scanning = WorkbookScanning.scanAsArr(workbook);
        XSSFWorkbook xwb = new XSSFWorkbook();
        Workbook wb = new SXSSFWorkbook(xwb, 100);
        int sheetIndex = 0;
        Sheet sheet = wb.createSheet("Extracted_"+sheetIndex);
        Date dateStart = new Date();
        System.out.println("Start time: "+dateStart);
        int row = 0;
        int count=0;
        while(scanning.hasNext()){
            if(count!=0 && count%1000==0){
                System.out.println(String.format("Working on %d at %s", count, new Date().toString()));
            }
            if(count%900000==0){
                sheetIndex++;
                sheet = wb.createSheet("Extracted_"+sheetIndex);
                row=0;
            }
            // System.out.println("Find next");
            String[] rowRecord = scanning.next();
            // System.out.println("Find complete");
            Row sheetRow = sheet.createRow(row);
            for(int i=0;i<rowRecord.length;i++) {
                // System.out.println("fill "+i);
                Cell cell = sheetRow.createCell(i);
                // System.out.println("Style 1");
                checkAndFillColor(((SXSSFWorkbook)wb).getXSSFWorkbook(), cell, i, 8, rowRecord);
                // System.out.println("Style 2");
                checkAndFillColor(((SXSSFWorkbook)wb).getXSSFWorkbook(), cell, i, 19, rowRecord);
                // System.out.println("Style 3");
                checkAndFillColor(((SXSSFWorkbook)wb).getXSSFWorkbook(), cell, i, 27, rowRecord);
                // System.out.println("fill start");
                cell.setCellValue(rowRecord[i]);
                // System.out.println("fill end");
            }
            System.out.println(String.format("[%d]Processing sheet: %s[%s]", count, rowRecord[0],rowRecord[3]));
            row++;
            count++;
        }
        workbook.close();
        workbook=null;
        try {
            Date dateStage2 = new Date();
            System.out.println("Start stage: "+dateStage2);
            String filePath = dest;
            if(new File(filePath).exists()) new File(filePath).delete();
            FileOutputStream outputStream = new FileOutputStream(filePath);
            wb.write(outputStream);
            wb.close();

            Date dateComplete = new Date();
            System.out.println("Start time: "+dateStart);
            System.out.println("Start stage: "+dateStage2);
            System.out.println("Completed: "+dateComplete);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
