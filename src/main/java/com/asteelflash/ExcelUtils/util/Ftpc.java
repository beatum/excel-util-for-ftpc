package com.asteelflash.ExcelUtils.util;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Vector;

import org.apache.poi.ss.usermodel.Workbook;

import com.asteelflash.ExcelUtils.exception.ExcelCommonException;

public class Ftpc {

	/**
	 * Name:createWorkSheet Description:creating new work sheet Author:Happy.He
	 * CreationTime:Feb 22, 2021 3:35:31 PM
	 */
	public static void exportExcel(String[] header, Vector vector, Integer pageSize, String path)
			throws ExcelCommonException {

		Integer pageCount = 1;
		Integer currentPage = 1;
		try {
			String fileExtension = path.substring(path.lastIndexOf("."));
			if (null == pageSize) {
				if (fileExtension.equalsIgnoreCase(".xls")) {
					// 65536
					pageSize = 30000;
				} else if (fileExtension.equalsIgnoreCase(".xlsx")) {
					pageSize = 1048576;
				} else {
					throw new ExcelCommonException("Incorrect file type!!");
				}
			}

			pageCount = (int) java.lang.Math.ceil((vector.size() / (double) pageSize));

			for (int i = 0; i < pageCount; i++) {
				Integer beginIndex = i * pageSize;
				Workbook workbook = Commons.createWorkBook(header, vector, beginIndex, pageSize, fileExtension);
				Integer index = path.lastIndexOf(".");
				String filePath = path.substring(0, index);
				FileOutputStream fileOutputStream = new FileOutputStream(
						filePath + "(" + (i + 1) + ")" + fileExtension);
				workbook.write(fileOutputStream);
				fileOutputStream.close();
			}
		} catch (Exception ex) {
			throw new ExcelCommonException(ex.getMessage());
		}

	}

	/**
	 * Name:importExcel Description: import excel  Author:Happy.He CreationTime:Feb 22, 2021
	 * 5:23:39 PM
	 */
	public static List<List<List<Object>>> importExcel(File file) throws ExcelCommonException {
		try {
			Workbook workbook = Commons.getWorkBookFromFile(file);
			return Commons.workSheet2List(workbook);
		} catch (Exception ex) {
			throw new ExcelCommonException(ex.getMessage());
		}

	}
	
	
	public static void main(String[] args) {
		Vector<Object> vector = new Vector<>();
		Vector<String> item = new Vector<>();
		item.add("123");
		item.add("456");
		vector.add(item.toArray());
		
		exportExcel ( new String[] {"happy","happy2"}, vector, null,"d:\\123.xlsx");
		
	}

}
