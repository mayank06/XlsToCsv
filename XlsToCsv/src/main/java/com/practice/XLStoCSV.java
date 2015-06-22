package com.practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import javax.imageio.stream.FileImageInputStream;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class XLStoCSV {

	public static void xls(File inputFile, File outputFile) throws IOException {

		StringBuffer data = new StringBuffer();
		try {
			FileOutputStream fos = new FileOutputStream(outputFile);

			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inputFile));

			HSSFSheet sheet = workbook.getSheetAt(0);
			Row row;
			Cell cell;
			Iterator<Row> rowIterator = sheet.iterator();
			while(rowIterator.hasNext()){
				row = rowIterator.next();

				Iterator<Cell> cellIterator = row.cellIterator();
				while(cellIterator.hasNext()) {

					cell = cellIterator.next();

					switch(cell.getCellType()) {

					case Cell.CELL_TYPE_BOOLEAN : 
						System.out.println(cell.getBooleanCellValue() + "\t\t");
						data.append(cell.getBooleanCellValue() + ",");
						break;
					case Cell.CELL_TYPE_NUMERIC :
						System.out.print(cell.getNumericCellValue() + "\t\t");
						data.append(cell.getNumericCellValue() + ",");
						break;
					case Cell.CELL_TYPE_STRING :
						System.out.print(cell.getStringCellValue() + "\t\t");
						data.append(cell.getStringCellValue() + ",");
						break;
					case Cell.CELL_TYPE_BLANK : 
						data.append("" + ",");
					default :
						data.append(cell + "");
					}
					data.append("\n");
				}
				System.out.println("");
			}
			fos.write(data.toString().getBytes());
			fos.close();

		}
		catch (FileNotFoundException e) {
			e.printStackTrace();
		}

	}

	public static void main(String[] args) throws IOException {
		File inputFile = new File("C:/Users/754098/Documents/TCS_GE_Confidential/test.xls");
		File outputFile = new File("C:/Users/754098/Documents/TCS_GE_Confidential/test.csv");
		xls(inputFile, outputFile);
	}
}

