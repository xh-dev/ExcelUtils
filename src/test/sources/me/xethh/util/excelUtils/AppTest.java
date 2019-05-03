package me.xethh.util.excelUtils;

import static org.junit.Assert.assertTrue;

import me.xethh.util.excelUtils.model.CellScanningModel;
import me.xethh.util.excelUtils.model.CellStyleScanningModel;
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
                switch (cell.getCellType()){
                    case BLANK:
                        model.setValue(null);
                        model.setCellType(CellScanningModel.CellType.Blank);
                        break;
                    case ERROR:
                        model.setValue(null);
                        model.setCellType(CellScanningModel.CellType.Error);
                        break;
                    case STRING:
                        model.setValue(cell.getStringCellValue());
                        model.setCellType(CellScanningModel.CellType.String);
                        break;
                    case NUMERIC:
                        if(HSSFDateUtil.isCellDateFormatted(cell)){
                            model.setValue(cell.getDateCellValue());
                            model.setCellType(CellScanningModel.CellType.Date);
                        }
                        else{
                            model.setValue(cell.getNumericCellValue());
                            model.setCellType(CellScanningModel.CellType.Decimal);
                        }
                        break;
                    case BOOLEAN:
                        model.setValue(cell.getBooleanCellValue());
                        model.setCellType(CellScanningModel.CellType.Boolean);
                        break;
                    case FORMULA:
                        model.setValue(cell.getCellFormula());
                        model.setCellType(CellScanningModel.CellType.Formula);
                        break;
                }

                System.out.println(model);
            }
        }
    }
}
