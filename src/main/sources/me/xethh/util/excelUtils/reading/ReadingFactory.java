package me.xethh.util.excelUtils.reading;

import me.xethh.util.excelUtils.common.ExcelReadValue;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
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
                checkAndFillColor(wb, cell, i, 8, rowRecord);
                checkAndFillColor(wb, cell, i, 19, rowRecord);
                checkAndFillColor(wb, cell, i, 27, rowRecord);
                // if(i==27 && rowRecord[i]!=null && !rowRecord[i].equals("") && rowRecord[i].split(",").length==3){
                //     byte[] fontColorByte = ExcelReadValue.toByteArr(rowRecord[i].split(","));
                //     if(!ExcelReadValue.isPureDark(rowRecord[i].split(","))){
                //         XSSFCellStyle style = wb.createCellStyle();
                //         style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                //         style.setFillForegroundColor(new XSSFColor(fontColorByte,wb.getStylesSource().getIndexedColors()));
                //         cell.setCellStyle(style);
                //     }
                // }
                // if(i==8 && rowRecord[8]!=null && !rowRecord[8].equals("") && rowRecord[8].split(",").length==4){
                //     byte[] fontColorByte = ExcelReadValue.toByteArr(rowRecord[i].split(","));
                //     if(!ExcelReadValue.isPureDark(rowRecord[i].split(","))){
                //         XSSFCellStyle style = wb.createCellStyle();
                //         style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                //         style.setFillForegroundColor(new XSSFColor(fontColorByte,wb.getStylesSource().getIndexedColors()));
                //         sheetRow.getCell(8).setCellStyle(style);
                //     }
                // }
                // if(i == 19 && rowRecord[19]!=null && !rowRecord[19].equals("") && rowRecord[19].split(",").length==4){
                //     byte[] fontColorByte = ExcelReadValue.toByteArr(rowRecord[i].split(","));
                //     if(!ExcelReadValue.isPureDark(rowRecord[i].split(","))){
                //         XSSFCellStyle style = wb.createCellStyle();
                //         style.setFillForegroundColor(new XSSFColor(fontColorByte,wb.getStylesSource().getIndexedColors()));
                //         style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                //         sheetRow.getCell(19).setCellStyle(style);
                //     }
                // }
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
