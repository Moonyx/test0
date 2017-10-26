package com.moonyx.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

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
		Workbook workbook = WorkbookFactory.create(stream);
		Map<String, List<List<Object>>> sheetDatas = new HashMap<String, List<List<Object>>>();
		for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
			Sheet sheet = workbook.getSheetAt(numSheet);
			if (sheet == null) {
				continue;
			}
			List<List<Object>> datas = new ArrayList<List<Object>>();
			for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
				Row row = sheet.getRow(rowNum);
				if (row == null) {
					continue;
				}
				List<Object> data = new ArrayList<Object>();
				boolean isRowAllNull = true;
				for (int cellNum = 0; cellNum <= row.getLastCellNum(); cellNum++) {
					Cell cell = null;
					if (isMergedRegion(sheet, rowNum, cellNum))
						cell = getMergedRegionCell(sheet, rowNum, cellNum);
					else
						cell = row.getCell(cellNum);
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
			sheetDatas.put(sheet.getSheetName(), datas);
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

	/**
	 * 判断指定的单元格是否是合并单元格
	 * 
	 * @param sheet
	 * @param row
	 *            行下标
	 * @param column
	 *            列下标
	 * @return
	 */
	private static boolean isMergedRegion(Sheet sheet, int row, int column) {
		int numMergedRegions = sheet.getNumMergedRegions();
		for (int i = 0; i < numMergedRegions; i++) {
			CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
			int firstColumn = cellRangeAddress.getFirstColumn();
			int lastColumn = cellRangeAddress.getLastColumn();
			int firstRow = cellRangeAddress.getFirstRow();
			int lastRow = cellRangeAddress.getLastRow();
			if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn)
				return true;
		}
		return false;
	}
	
	/**
	 * 获取合并单元格的值
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	private static Cell getMergedRegionCell(Sheet sheet, int row, int column) {
		int numMergedRegions = sheet.getNumMergedRegions();
		for (int i = 0; i < numMergedRegions; i++) {
			CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
			int firstColumn = cellRangeAddress.getFirstColumn();
			int lastColumn = cellRangeAddress.getLastColumn();
			int firstRow = cellRangeAddress.getFirstRow();
			int lastRow = cellRangeAddress.getLastRow();
			if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn) {
				Row r = sheet.getRow(firstRow);
				return r.getCell(firstColumn);
			}
		}
		return null;
	}
	
	public static void main(String[] args) {
//		File file = new File("C:\\Users\\user\\Desktop\\CKSzz\\LT\\進口-北上資料-new\\Open PO(0920).xlsx");
		File file = new File("C:\\Users\\user\\Desktop\\CKSzz\\LT\\進口-北上資料-new\\Open PO-1 - 副本.xlsx");
		try {
			readExcelSheetDatasForBigGrid(new FileInputStream(file), "Open PO(0920).xlsx");
//			System.out.println("s");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * 获取所有excel数据（大数据量）
	 * 
	 * @author lxh
	 * @param
	 * @return
	 */
	public static Map<String, List<List<Object>>> readExcelSheetDatasForBigGrid(InputStream stream, String fileName) throws Exception {
		if (stream == null || StringUtils.isEmpty(fileName))
			throw new Exception("未知文件格式。");
		if (!fileName.endsWith(".xlsx")) {
			throw new Exception("只支持.xlsx格式。");
		}
		Map<String, List<List<Object>>> sheetDatas = new LinkedHashMap<String, List<List<Object>>>();
		List<List<Object>> rowList = new ArrayList<List<Object>>();
		
		OPCPackage opcPackage = OPCPackage.open(stream);
		XSSFReader xssfReader = new XSSFReader(opcPackage);
		SharedStringsTable sharedStringsTable = xssfReader.getSharedStringsTable();
		StylesTable stylesTable = xssfReader.getStylesTable();
		XMLReader parser = fetchSheetParser(sharedStringsTable, stylesTable, rowList);

		Iterator<InputStream> sheetsDataInputStreamIterator = xssfReader.getSheetsData();
		if (sheetsDataInputStreamIterator.hasNext()) {
			InputStream sheetsDataInputStream = sheetsDataInputStreamIterator.next();
			InputSource sheetsInputSource = new InputSource(sheetsDataInputStream);
			parser.parse(sheetsInputSource);
			sheetsDataInputStream.close();
		}
		sheetDatas.put("Sheet1", rowList);// 暂时支持单sheet
		return sheetDatas;
	}

	public static XMLReader fetchSheetParser(SharedStringsTable sst, StylesTable stylesTable, List<List<Object>> rowList) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader("com.sun.org.apache.xerces.internal.parsers.SAXParser");//org.apache.xerces.parsers.SAXParser
		ContentHandler handler = new SheetHandler(sst, stylesTable, rowList);
		parser.setContentHandler(handler);
		return parser;
	}

	/**
	 * See org.xml.sax.helpers.DefaultHandler javadocs
	 */
	private static class SheetHandler extends DefaultHandler {
		
		private SharedStringsTable sst;
		private StylesTable stylesTable;
		private List<List<Object>> rowList;
		private List<Object> cellList;
		
		private String lastContents;
		private String lastName;
		private String cellStyle;
		private boolean nextIsString;

		private SheetHandler(SharedStringsTable sst, StylesTable stylesTable, List<List<Object>> rowList) {
			this.sst = sst;
			this.stylesTable = stylesTable;
			this.rowList = rowList;
			this.cellList = new ArrayList<Object>();
		}

		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
			if ("c".equals(name)) {
				String cellType = attributes.getValue("t");
				cellStyle = attributes.getValue("s");
				if (cellType != null && cellType.equals("s")) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
			}
			lastContents = "";
		}

		public void endElement(String uri, String localName, String name) throws SAXException {
			if (nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				nextIsString = false;
			}
			System.out.println(name + "--" + lastContents);
			if ("v".equals(name)) {
				try {
					if (cellStyle != null) {
						int styleIndex = Integer.parseInt(cellStyle);
						XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
						int dataFormat = style.getDataFormat();
						String dataFormatString = style.getDataFormatString();
						Pattern dateCn = Pattern.compile(".*[a-z]{2,4}\"年\"m\"月\"(d\"日\")?.*");
						if ((dateCn.matcher(dataFormatString).matches() || DateUtil.isADateFormat(dataFormat, dataFormatString)) 
								&& StringUtils.isNumeric(lastContents))
							cellList.add(HSSFDateUtil.getJavaDate(Double.valueOf(lastContents)));
						else
							cellList.add(lastContents);
					} else {
						cellList.add(lastContents);
					}
				} catch (Exception e) {
					e.printStackTrace();
					cellList.add(lastContents);
				}
			} else if ("c".equals(name) && name.equals(this.lastName)) {
				cellList.add(null);
			}
			if ("row".equals(name)) {
				rowList.add(cellList);
				cellList = new ArrayList<Object>();
			}
			this.lastName = name;
		}

		public void characters(char[] ch, int start, int length) throws SAXException {
			lastContents += new String(ch, start, length);
		}
	}

}
