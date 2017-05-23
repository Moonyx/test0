package com.moonyx.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPOIUtils {
	
	private static Logger logger = LogManager.getLogger(ExcelPOIUtils.class);
	
	public static Map<String, List<List<Object>>> readExcelSheetDatas(File file) throws Exception {
		String fileName = file.getName().toLowerCase();
		String excelType = "07";
		if (fileName.endsWith(".xls"))
			excelType = "03";
		else if (fileName.endsWith(".xlsx"))
			excelType = "07";
		else
			throw new Exception("error-文件格式不对，请选取.xls或.xlsx文件！");
		return readExcelSheetDatas(new FileInputStream(file), excelType);
	}
	
	public static Map<String, List<List<Object>>> readExcelSheetDatas(InputStream stream, String excelType) throws Exception {
		Workbook workbook = null;
		try {
			if ("03".equals(excelType)) {
				workbook = new HSSFWorkbook(stream);
			} else {
				workbook = new XSSFWorkbook(stream);
			}
			Map<String, List<List<Object>>> sheetDatas = new HashMap<String, List<List<Object>>>();
			for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
				Sheet hssfSheet = workbook.getSheetAt(numSheet);
				if (hssfSheet == null) {
					continue;
				}
				List<List<Object>> datas = new ArrayList<List<Object>>();
				for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
					Row hssfRow = hssfSheet.getRow(rowNum);
					if (hssfRow == null) {
						continue;
					}
					List<Object> data = new ArrayList<Object>();
					boolean isRowAllNull = true;
					for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++) {
						Cell cell = hssfRow.getCell(cellNum);
						if (cell == null) {
							data.add(null);
							continue;
						}
						Object cellValue = getCellValue(cell);
						data.add(cellValue);
						if (isRowAllNull && cellValue != null 
								&& !"".equals(cellValue.toString().trim()) && !"NULL".equals(cellValue.toString()))
							isRowAllNull = false;
					}
					if (!isRowAllNull)
						datas.add(data);
				}
				sheetDatas.put(hssfSheet.getSheetName(), datas);
			}
			return sheetDatas;
		} catch (Exception e) {
			logger.error(e);
			throw e;
		} finally {
			if (workbook != null)
				workbook.close();
		}
	}
	
	private static Object getCellValue(Cell cell) {
		Object result = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			result = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (isCellInternalDateFormatted(cell))
				result = cell.getDateCellValue();
			else
				result = cell.getNumericCellValue();
			
			break;
		case Cell.CELL_TYPE_FORMULA:
			try {
				result = cell.getStringCellValue();
			} catch (IllegalStateException e) {
				result = cell.getNumericCellValue();
			}
			break;
		case Cell.CELL_TYPE_ERROR:
			result = cell.getErrorCellValue();
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			result = cell.getBooleanCellValue();
			break;
		default:
			result = "NULL";
			break;
		}
		return result;
	}
	
	private static boolean isCellInternalDateFormatted(Cell cell){
		if (cell == null)
			return false;
		boolean bDate = false;
		double d = cell.getNumericCellValue();
		if ( DateUtil.isValidExcelDate(d) ) {
			CellStyle style = cell.getCellStyle();
			if(style==null) return false;
			int i = style.getDataFormat();
			String f = style.getDataFormatString();
			Pattern dateCn = Pattern.compile(".*[a-z]{2,4}\"年\"m\"月\"(d\"日\")?.*");
			if(dateCn.matcher(f).matches())
				return true;
			bDate = DateUtil.isADateFormat(i, f);
		}
		return bDate;
	}

}
