package me.xethh.util.excelUtils;

import static org.junit.Assert.assertTrue;

import me.xethh.util.excelUtils.common.ExcelReadValue;
import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.util.excelUtils.model.CellStyleScanningModel;
import me.xethh.util.excelUtils.reading.WorkbookScanning;
import me.xethh.utils.wrapper.Tuple2;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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

    public static void main(String[] args) throws IOException {
        InputStream is = new FileInputStream(new File("./src/test/resources/TestingBase.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        // Iterator<CellScanningModel> scanning = WorkbookScanning.scan(workbook);
        // while(scanning.hasNext()){
        //     System.out.println(scanning.next());
        // }
        Iterator<String[]> scanning = WorkbookScanning.scanAsArr(workbook);
        while(scanning.hasNext()){
            for(String s:scanning.next())
                System.out.print(s+", ");
            System.out.println();
        }
        // FileOutputStream os = new FileOutputStream(new File("./src/test/resources/TestingBase_revised.xlsx"));
        
    }
}
