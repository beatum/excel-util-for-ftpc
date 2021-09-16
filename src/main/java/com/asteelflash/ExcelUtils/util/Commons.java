package com.asteelflash.ExcelUtils.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;
import java.util.Vector;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.asteelflash.ExcelUtils.exception.ExcelCommonException;

/**
 * @author Happy.He
 *
 */
public class Commons {
	public Commons() {
	}

	/**
	 * Name:getFileExtension Description:get file extension Author:Happy.He
	 * CreationTime:Feb 22, 2021 1:30:21 PM
	 */
	public static String getFileExtension(File file) throws IOException {
		String fileName = file.getName();
		Integer endIndex = fileName.lastIndexOf(".");
		return fileName.substring(endIndex);
	}

	/**
	 * Name:getWorkBookFromFile Description:getinng excel workbook from a file
	 * Author:Happy.He CreationTime:Feb 22, 2021 1:21:08 PM
	 */
	public static Workbook getWorkBookFromFile(File file) throws ExcelCommonException {
		try {
			Workbook workbook = null;
			String fileExtension = getFileExtension(file);
			FileInputStream fileInputStream = new FileInputStream(file);
			// excel 97-2003 .xls
			if (fileExtension.equalsIgnoreCase(".xls")) {
				workbook = new HSSFWorkbook(fileInputStream);
			}
			// excel 2007+ .xlsx
			if (fileExtension.equalsIgnoreCase(".xlsx")) {
				workbook = new XSSFWorkbook(fileInputStream);
			}

			if (null == workbook) {
				throw new ExcelCommonException("Incorrect file type!!");
			} else {
				return workbook;
			}

		} catch (Exception ex) {
			throw new ExcelCommonException(ex.getMessage());
		}

	}

	/**
	 * Name:getWorkSheets Description:geting excel work sheets Author:Happy.He
	 * CreationTime:Feb 22, 2021 1:54:01 PM
	 */
	public static List<Sheet> getWorkSheets(Workbook workbook) throws ExcelCommonException {
		List<Sheet> sheets = null;
		try {
			Integer count = workbook.getNumberOfSheets();
			sheets = new LinkedList<>();
			for (int i = 0; i < count; i++) {
				Sheet sheet = workbook.getSheetAt(i);
				sheets.add(sheet);
			}
		} catch (Exception ex) {
			throw new ExcelCommonException(ex.getMessage());
		}
		return sheets;
	}

	/**
	 * Name:workSheet2List Description: converting a excel work sheet to list
	 * Author:Happy.He CreationTime:Feb 22, 2021 2:09:54 PM
	 */
	public static List<List<List<Object>>> workSheet2List(Workbook workbook) throws ExcelCommonException {
		List<List<List<Object>>> workSheets = null;
		try {
			workSheets = new LinkedList<>();
			List<Sheet> sheets = getWorkSheets(workbook);
			for (int i = 0; i < sheets.size(); i++) {
				List<List<Object>> item = sheet2List(sheets.get(i));
				workSheets.add(item);
			}
		} catch (Exception ex) {
			new ExcelCommonException(ex.getMessage());
		}
		return workSheets;
	}

	/**
	 * Name:sheet2List Description: data sheet to list Author:Happy.He
	 * CreationTime:Feb 22, 2021 2:53:11 PM
	 */
	private static List<List<Object>> sheet2List(Sheet sheet) throws ExcelCommonException {
		List<List<Object>> columns = null;
		try {
			DataFormatter formatter = new DataFormatter();
			columns = new LinkedList<>();
			for (Row row : sheet) {
				List<Object> rowItems = new LinkedList<>();
				for (Cell cell : row) {
					// CellReference cellRef = new CellReference(row.getRowNum(),
					// cell.getColumnIndex());
					// System.out.print(cellRef.formatAsString());
					// System.out.print(" - ");
					// String text = formatter.formatCellValue(cell);
					// System.out.println(text);
					// Alternatively, get the value and format it yourself
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						// System.out.println(cell.getRichStringCellValue().getString());
						rowItems.add(cell.getRichStringCellValue().toString());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						if (DateUtil.isCellDateFormatted(cell)) {
							// System.out.println(cell.getDateCellValue());
							rowItems.add(cell.getDateCellValue());
						} else {
							// System.out.println(cell.getNumericCellValue());
							rowItems.add(cell.getNumericCellValue());
						}
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						// System.out.println(cell.getBooleanCellValue());
						rowItems.add(cell.getBooleanCellValue());
						break;
					case Cell.CELL_TYPE_FORMULA:
						System.out.println(cell.getCellFormula());
						rowItems.add(cell.getCellFormula());
						break;
					case Cell.CELL_TYPE_BLANK:
						// System.out.println();
						rowItems.add(null);
						break;
					default:
						// System.out.println();
						rowItems.add(null);
					}
				}
				columns.add(rowItems);
			}

		} catch (Exception ex) {
			throw new ExcelCommonException(ex.getMessage());
		}
		return columns;
	}

	/**
	 * Name:createWorkBook Description: creating a new workbook Author:Happy.He
	 * CreationTime:Feb 22, 2021 4:56:52 PM
	 */
	public static Workbook createWorkBook(String[] header, Vector vector, Integer startIndex, Integer pageSize,
			String fileExtension) throws ExcelCommonException {
		Workbook workbook = null;
		if (fileExtension.equalsIgnoreCase(".xlsx")) {
			workbook = new XSSFWorkbook();
		} else if (fileExtension.equalsIgnoreCase(".xls")) {
			workbook = new HSSFWorkbook();
		} else {
			throw new ExcelCommonException("Incorrect file extension!!");
		}

		Sheet sheet = workbook.createSheet("Sheet1");
		// create header start
		Row row_header = sheet.createRow(0);
		for (int i = 0; i < header.length; i++) {
			Cell header_cell = row_header.createCell(i);
			String element = header[i];
			if (element == null) {
				element = "";
			}
			header_cell.setCellValue(element);
		}
		// create header end
		Integer rowIndex = 1;
		for (int i = startIndex; i < vector.size(); i++) {
			if (rowIndex > pageSize) {
				break;
			}
			Row row = sheet.createRow(rowIndex);
			Object[] elements = (Object[]) vector.elementAt(i);
			for (int j = 0; j < elements.length; j++) {
				Cell cell = row.createCell(j);
				Object element = elements[j];
				if (element == null) {
					element = "";
				}
				cell.setCellValue(element.toString());
			}
			rowIndex++;
		}
		return workbook;
	}

	public static void main(String[] args) throws Exception {
		String filePath = "D:\\OrderList.xlsx";
		Workbook workbook = getWorkBookFromFile(new File(filePath));

		Object result = workSheet2List(workbook);
		System.out.println("");

	}

}
