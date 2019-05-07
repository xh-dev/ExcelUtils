package me.xethh.util.excelUtils;

import static org.junit.Assert.assertTrue;

import me.xethh.util.excelUtils.common.ExcelReadValue;
import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.util.excelUtils.model.CellStyleScanningModel;
import me.xethh.util.excelUtils.reading.ReadingFactory;
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

    public static void main(String[] args) throws IOException {
        ReadingFactory.scanExcel("./src/test/resources/TestingBase.xlsx","./target/TestingBase_revised.xlsx");
    }
}
