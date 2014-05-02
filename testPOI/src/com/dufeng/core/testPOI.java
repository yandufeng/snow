package com.dufeng.core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * http://viralpatel.net/blogs/java-read-write-excel-file-apache-poi/
 * 
 * @author yandufeng
 * 
 */
public class testPOI {

	public static void main(String[] args) throws IOException {
		print();

		// create();

		// update();

		// formula();
	}

	private static void print() {
		try {

			FileInputStream file = new FileInputStream(
					new File("F:\\test.xlsx"));

			// Get the workbook instance for XLS file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				// For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						System.out.print(cell.getBooleanCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + "\t\t");
						break;
					}
				}
				System.out.println("");
			}

			List<CellRangeAddress> regionsList = new ArrayList<CellRangeAddress>();
			for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
				regionsList.add(sheet.getMergedRegion(i));
			}
			int lastRowNum = sheet.getLastRowNum();
			
			int maxColumnNum = sheet.getRow(0).getLastCellNum();
			int columnNum = 0;
			Row row = null;
			for(int i = 1; i < lastRowNum; i++) {
				row = sheet.getRow(i);
				if(row != null) {
					columnNum = row.getLastCellNum();
					if(maxColumnNum < columnNum) {
						maxColumnNum = columnNum;
					}
				}
			}
			System.out.println("regionsList =" + regionsList);
			for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
				Row r = sheet.getRow(rowNum);
				if (r != null) {
					// Row r = rowIterator.next();
					int lastColumn = r.getLastCellNum();
					// System.out.print("r.getLastCellNum = " + lastColumn);
					for (int cn = 0; cn < maxColumnNum; cn++) {
						Cell cell = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);
						for (CellRangeAddress region : regionsList) {
							if (region.isInRange(rowNum, cn)) {
								int rowIndex = region.getFirstRow();
								int colIndex = region.getFirstColumn();
								cell = sheet.getRow(rowIndex).getCell(colIndex);
								break;
							}
						}
						if (cell == null) {
							System.out.print("Hello\t\t");
						} else {
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_BOOLEAN:
								System.out.print(cell.getBooleanCellValue()
										+ "\t\t");
								break;
							case Cell.CELL_TYPE_NUMERIC:
								System.out.print(cell.getNumericCellValue()
										+ "\t\t");
								break;
							case Cell.CELL_TYPE_STRING:
								System.out.print(cell.getStringCellValue()
										+ "\t\t");
								break;
							}
						}
					}
					System.out.println("");
				}else {
					for(int i = 0; i < maxColumnNum; i++) {
						System.out.print("hello\t\t");
					}
					System.out.println("");
				}
			}
			file.close();
			FileOutputStream out = new FileOutputStream(
					new File("F:\\test.xls"));
			workbook.write(out);
			out.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void create() {

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Sample sheet");

		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("3", new Object[] { "Emp No.", "Name", "Salary" });
		data.put("2", new Object[] { 1d, "John", 1500000d });
		data.put("1", new Object[] { 2d, "Sam", 800000d });
		data.put("4", new Object[] { 3d, "Dean", 700000d });

		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof Date)
					cell.setCellValue((Date) obj);
				else if (obj instanceof Boolean)
					cell.setCellValue((Boolean) obj);
				else if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Double)
					cell.setCellValue((Double) obj);
			}
		}

		try {
			FileOutputStream out = new FileOutputStream(new File("F:\\new.xls"));
			workbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static void update() {
		try {
			FileInputStream file = new FileInputStream(new File("F:\\new.xls"));

			HSSFWorkbook workbook = new HSSFWorkbook(file);
			HSSFSheet sheet = workbook.getSheetAt(0);
			Cell cell = null;

			// Update the value of cell
			cell = sheet.getRow(1).getCell(2);
			cell.setCellValue(cell.getNumericCellValue() * 2);
			cell = sheet.getRow(2).getCell(2);
			cell.setCellValue(cell.getNumericCellValue() * 2);
			cell = sheet.getRow(3).getCell(2);
			cell.setCellValue(cell.getNumericCellValue() * 2);

			file.close();

			FileOutputStream outFile = new FileOutputStream(new File(
					"F:\\new.xls"));
			workbook.write(outFile);
			outFile.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static void formula() {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Calculate Simple Interest");

		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("Pricipal Amount (P)");
		header.createCell(1).setCellValue("Rate of Interest (r)");
		header.createCell(2).setCellValue("Tenure (t)");
		header.createCell(3).setCellValue("Interest (P r t)");

		Row dataRow = sheet.createRow(1);
		dataRow.createCell(0).setCellValue(14500d);
		dataRow.createCell(1).setCellValue(9.25);
		dataRow.createCell(2).setCellValue(3d);
		dataRow.createCell(3).setCellFormula("A2*B2*C2");

		try {
			FileOutputStream out = new FileOutputStream(new File(
					"C:\\formula.xls"));
			workbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
