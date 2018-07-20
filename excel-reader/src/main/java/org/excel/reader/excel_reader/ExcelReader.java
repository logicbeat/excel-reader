package org.excel.reader.excel_reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Program to read an excel file and print the contents of the file in an output log
 */
public class ExcelReader 
{
	private static final String FILE_NAME = "D:/temp/Book1.xlsx";
	private static final String OUTPUT_FILE_NAME = "D:/temp/output.log";

    public static void main( String[] args )
    {
        try (FileInputStream fis = new FileInputStream(new File(FILE_NAME));
        		FileOutputStream fos = new FileOutputStream(new File(OUTPUT_FILE_NAME));
        		PrintWriter writer = new PrintWriter(fos);
        		Workbook book = new XSSFWorkbook(fis)) {
			Sheet currentSheet = book.getSheetAt(0);
			Iterator<Row> rows = currentSheet.iterator();
			while (rows.hasNext()) {
				Row currentRow = rows.next();
				Iterator<Cell> cellIterator = currentRow.cellIterator();
				while(cellIterator.hasNext()){
					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        writer.write(currentCell.getStringCellValue() + " ");
                    } else if(currentCell.getCellTypeEnum() == CellType.NUMERIC){
                    	writer.write(currentCell.getNumericCellValue() + " ");
                    }
				}
				writer.write("\n");
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
        
    }
}
