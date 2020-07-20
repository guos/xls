package com.uonow.finder.xls;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	
	
	public static Workbook getWorkbook(String fileName) {
		Workbook wb = null;
		if (fileName == null) {
			return null;
		}
		String extString = fileName.substring(fileName.lastIndexOf("."));
		InputStream is = null;
		try {
			is = new FileInputStream(fileName);
			if (".xls".equals(extString)) {
				return wb = new HSSFWorkbook(is);
			} else if (".xlsx".equals(extString)) {
				return wb = new XSSFWorkbook(is);
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return wb;

	}

	public static List<Map<String, String>> readExceList(String fileName, String columns[]) {
		Sheet sheet = null;
		Row row = null;
		Row rowHeader = null;
		List<Map<String, String>> list = null;
		String cellData = null;

		try {
			Workbook wb = getWorkbook(fileName); // 获得excel文件对象workbook

			if (wb != null) {
				// 用来存放表中数据
				list = new ArrayList<Map<String, String>>();
				// 获取第一个sheet
				sheet = wb.getSheetAt(0);
				// 获取最大行数
				int rownum = sheet.getPhysicalNumberOfRows();
				// 获取第2行
				rowHeader = sheet.getRow(1);
				row = sheet.getRow(1);
				// 获取最大列数
				int colnum = row.getPhysicalNumberOfCells();
				//System.out.println("colum:" + colnum);
				for (int i = 2; i < rownum; i++) {
					Map<String, String> map = new LinkedHashMap<String, String>();
					row = sheet.getRow(i);
					if (row != null) {
						for (int j = 0; j < colnum; j++) {
							if (columns[j].equals(getCellFormatValue(rowHeader.getCell(j)))) {
								cellData = (String) getCellFormatValue(row.getCell(j));
								map.put(columns[j], cellData);
							}
						}
					} else {
						break;
					}
					list.add(map);
				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return list;
	}

	public static String getCellFormatValue(Cell cell) {
		String cellValue = "";
		if (cell != null) {
			// 判断cell类型
			switch (cell.getCellType()) {
			case NUMERIC: {
				cellValue = String.valueOf(cell.getNumericCellValue());
				break;
			}
			case STRING: {
				cellValue = cell.getRichStringCellValue().getString();
				break;
			}
			default:
				cellValue = "";
			}
		}
		return cellValue;

	}

}
