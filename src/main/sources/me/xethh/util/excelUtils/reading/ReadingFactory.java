package me.xethh.util.excelUtils.reading;

import me.xethh.util.excelUtils.common.ColorUtils;
import me.xethh.util.excelUtils.common.ExcelReadValue;
import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.utils.wrapper.Tuple3;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ReadingFactory {
    private static void checkAndFillColor(XSSFWorkbook wb, Cell cell, int nowIndex, int index, String[] rowRecord){
        Pattern pattern = Pattern.compile("TempCTColor\\{a=(\\d+), r=(\\d+), g=(\\d+), b=(\\d+), tint=([\\-0-9\\.]+)\\}");

        if(nowIndex==index && rowRecord[index]!=null && !rowRecord[index].equals("")){
            Matcher matcher = pattern.matcher(rowRecord[index]);
            if(matcher.matches()){
                double tint = Double.valueOf(matcher.group(5));
                int a = Integer.parseInt(matcher.group(1));
                int r = ColorUtils.applyTint(Integer.parseInt(matcher.group(2)),tint);
                int g = ColorUtils.applyTint(Integer.parseInt(matcher.group(3)),tint);
                int b = ColorUtils.applyTint(Integer.parseInt(matcher.group(4)),tint);

                byte[] bytes = new byte[]{(byte) a, (byte) r, (byte) g, (byte) b};

                XSSFColor color = new XSSFColor(bytes, wb.getStylesSource().getIndexedColors());
                // color.setTint(tint);
                XSSFCellStyle style = wb.createCellStyle();
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setFillForegroundColor(color);
                cell.setCellStyle(style);
            }
        }
    }
    public static void scanExcel(InputStream is, String dest, Iterator<CellScanningModel> cellModelList){

        XSSFWorkbook xwb = new XSSFWorkbook();
        Workbook wb = new SXSSFWorkbook(xwb, 100);
        Sheet sheet = wb.createSheet("Extracted");
        Date dateStart = new Date();
        System.out.println("Start time: "+dateStart);
        sheet.createRow(0);
        for(int i=0; i < CellScanningModel.toStringArrHeader().length;i++)
            sheet.getRow(0).createCell(i).setCellValue(CellScanningModel.toStringArrHeader()[i]);

        int row = 1;
        while(cellModelList.hasNext()){
            // System.out.println("Find next");
            CellScanningModel next = cellModelList.next();
            String[] rowRecord = next.toStringArr();
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
            row++;
        }
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

    public static void scanExcel(String source, String dest, List<Tuple3<String, String,String>> areaList) throws IOException {
        try(
                InputStream is = new FileInputStream(new File(source));
                XSSFWorkbook workbook = new XSSFWorkbook(is);
                )
        {
            Iterator<CellScanningModel> scanning = WorkbookScanning.scan(workbook, areaList);
            scanExcel(is, dest, scanning);
        }
        finally{

            System.out.println("Completed extracting");
        }
    }
    public static void scanExcel(String source, String dest) throws IOException {
        try(
                InputStream is = new FileInputStream(new File(source));
                XSSFWorkbook workbook = new XSSFWorkbook(is);
                )
        {
            Iterator<CellScanningModel> scanning = WorkbookScanning.scan(workbook);
            scanExcel(is, dest, scanning);

        }
    }
}
