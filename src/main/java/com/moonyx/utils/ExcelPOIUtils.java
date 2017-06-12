package com.moonyx.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.lang.StringUtils;
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

	private static final SimpleDateFormat defaultFormat = new SimpleDateFormat("yyyy-MM-dd");
	
	public static Map<String, List<List<Object>>> readExcelSheetDatas(File file) throws Exception {
		String fileName = file.getName().toLowerCase();
		if (!fileName.endsWith(".xls") && !fileName.endsWith(".xlsx"))
			throw new Exception("error-文件格式不对，请选取.xls或.xlsx文件！");
		return readExcelSheetDatas(new FileInputStream(file), fileName);
	}

	public static Map<String, List<List<Object>>> readExcelSheetDatas(InputStream stream, String fileName) throws Exception {
		if (stream == null || StringUtils.isEmpty(fileName))
			throw new Exception("未知文件格式。");;
		Workbook workbook = null;
		if (fileName.endsWith(".xls")) {
			workbook = new HSSFWorkbook(stream);
		} else if (fileName.endsWith(".xlsx")) {
			workbook = new XSSFWorkbook(stream);
		} else {
			throw new Exception("只支持.xls / .xlsx格式。");
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
							&& !"".equals(cellValue.toString().trim())
							&& !"NULL".equals(cellValue.toString()))
						isRowAllNull = false;
				}
				if (!isRowAllNull)
					datas.add(data);
			}
			sheetDatas.put(hssfSheet.getSheetName(), datas);
		}
		return sheetDatas;
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

	
	public static Object getCellObject(List<List<Object>> sheetData, int row, int col) {
		if (sheetData.size() <= row)
			return null;
		List<Object> rowData = sheetData.get(row);
		if (rowData.size() <= col || col < 0)
			return null;
		Object cellObj = rowData.get(col);
		if (cellObj == null || "".equals(cellObj.toString()))
			return null;
		return cellObj;
	}

	/***
	 * 转字符串
	 * 
	 * @param 
	 * @return
	 */
	public static String objectToSting(Object obj) {
		String ret = null;
		if (obj == null || "".equals(obj.toString())) {
			return null;
		}
		if (obj instanceof String) {
			ret = ((String) obj).trim();
		} else if (obj instanceof Double) {
			ret = String.valueOf((Double) obj);
		} else if (obj instanceof Character) {
			ret = String.valueOf((Character) obj);
		} else if (obj instanceof Float) {
			ret = String.valueOf((Float) obj);
		} else if (obj instanceof Integer) {
			ret = String.valueOf((Integer) obj);
		} else if (obj instanceof Long) {
			ret = String.valueOf((Long) obj);
		} else if (obj instanceof Date) {
			ret = defaultFormat.format((Date) obj);
		} else {
			throw new RuntimeException("字符串类型不正确");
		}
		if (!(obj instanceof String) && !(obj instanceof Date) && !(obj instanceof Character)) {
			ret = objectToBigDecimal(obj).toPlainString();
		}
		return ret;
	}

	/***
	 * 转数字BigDecimal
	 * 
	 * @param o
	 * @return
	 */
	public static BigDecimal objectToBigDecimal(Object o) {
		BigDecimal ret = null;
		if (o == null || "".equals(o.toString())) {
			ret = new BigDecimal(0);
			return ret;
		}
		if (o instanceof Double) {
			ret = new BigDecimal((Double) o);
		} else if (o instanceof String) {
			try {
				ret = new BigDecimal((String) o);
			} catch (Exception e) {
				throw new RuntimeException("数字类型不正确");
			}
		} else if (o instanceof Character) {
			try {
				ret = new BigDecimal((Character) o);
			} catch (Exception e) {
				throw new RuntimeException("数字类型不正确");
			}
		} else if (o instanceof Float) {
			ret = new BigDecimal((Float) o);
		} else if (o instanceof Integer) {
			ret = new BigDecimal((Integer) o);
		} else if (o instanceof Long) {
			ret = new BigDecimal((Long) o);
		} else {
			throw new RuntimeException("数字类型不正确");
		}
		return ret;
	}
	
	/***
	 * 转数字Long
	 * 
	 * @param o
	 * @return
	 */
	public static Long objectToLong(Object o) {
		Long ret = 0l;
		if (o == null || "".equals(o.toString())) {
			return ret;
		}
		try {
			BigDecimal b = objectToBigDecimal(o);
			ret = b.longValue();
		} catch (Exception e) {
			return 0l;
		}
		return ret;
	}

	/***
	 * 转日期
	 * 
	 * @param o
	 * @return
	 */
	public static Date objectToDate(Object o, String formartStr) {
		if (o == null || "".equals(o.toString())) {
			return null;
		}
		Date ret = null;
		if (o instanceof Date) {
			ret = (Date) o;
		} else {
			try {
				SimpleDateFormat formatter = new SimpleDateFormat(formartStr);
				ret = formatter.parse(objectToSting(o));
			} catch (Exception e) {
				throw new RuntimeException("日期格式錯誤，正確格式:" + formartStr + "！");
			}
		}
		return ret;
	}

}
