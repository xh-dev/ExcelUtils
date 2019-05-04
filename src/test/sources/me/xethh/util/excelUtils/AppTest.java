package me.xethh.util.excelUtils;

import static org.junit.Assert.assertTrue;

import me.xethh.util.excelUtils.common.ExcelReadValue;
import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.util.excelUtils.model.CellStyleScanningModel;
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
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIt = sheet.rowIterator();
        while(rowIt.hasNext()){
            Iterator<Cell> cellIt = rowIt.next().cellIterator();
            while (cellIt.hasNext()){
                Cell cell = cellIt.next();
                CellScanningModel model = new CellScanningModel();
                model.setActRow(cell.getRowIndex()+1);
                model.setActCol(cell.getColumnIndex());
                model.setCellStyle(new CellStyleScanningModel(workbook,cell.getCellStyle()));
                model.setCellStr("");
                Tuple2<CellScanningModel.CellType, Object> value = ExcelReadValue.read(cell);
                model.setValue(value.getV2());
                model.setCellType(value.getV1());
                System.out.println(model);
            }
        }
    }
}
