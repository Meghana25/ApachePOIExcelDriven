package apache.poi.api;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class RetrieveExcelData {

    public ArrayList<String> getData(String testcasename) throws IOException {
        ArrayList<String> data = new ArrayList<>();
        FileInputStream fis = new FileInputStream(System.getProperty("user.dir")+"/src/main/resources/SampleData.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
        for(int i=0;i<xssfWorkbook.getNumberOfSheets();i++)
        {
            if(xssfWorkbook.getSheetName(i).equalsIgnoreCase("testdata"))
            {
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(i);
                Iterator<Row> rowIterator = xssfSheet.rowIterator();
                Row firstRow = rowIterator.next();
                Iterator<Cell> cellIterator = firstRow.cellIterator();
                int columnIndex=0;
                while (cellIterator.hasNext())
                {
                    Cell cellValue = cellIterator.next();
                    if(cellValue.getStringCellValue().equalsIgnoreCase("TestCases"))
                    {
                        columnIndex = cellValue.getColumnIndex();
                        break;
                    }
                }
                while (rowIterator.hasNext())
                {
                    Row testCasesColumn = rowIterator.next();
                    if(testCasesColumn.getCell(columnIndex).getStringCellValue().equalsIgnoreCase("Purchase"))
                    {
                        Iterator<Cell> testCaseRowIterator = testCasesColumn.cellIterator();
                        while (testCaseRowIterator.hasNext())
                        {
                            Cell cellType = testCaseRowIterator.next();
                            if(cellType.getCellType() == CellType.STRING) {
                                data.add(cellType.getStringCellValue());
                            }
                            else{
                                data.add(cellType.getDateCellValue().toString());
                            }
                        }

                    }
                }
            }
        }
        return data;
    }
}
