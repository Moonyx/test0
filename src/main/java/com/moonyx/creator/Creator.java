package com.moonyx.creator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import com.moonyx.utils.ExcelPOIUtils;
import com.moonyx.utils.StringUtils;

public class Creator {
	
	public static String path = "D:\\work\\create\\design\\";
	public static String modelPath = "D:\\work\\create\\design\\model\\";
	public static String targetPath = "D:\\work\\create\\design\\target\\";
	
	public static SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
	
	public static String tableName;
	public static String classFullName;
	public static String objectName;
	public static String className;
	
	public static void main(String[] args) {
		File file = new File(path + "designDB.xlsx");
		try {
			Map<String, List<List<Object>>> datas = ExcelPOIUtils.readExcelSheetDatas(new FileInputStream(file), file.getName());
			parseTableName(datas);
			parseClassFullName(datas);
			parseObjectName();
			parseClassName();
			String targetFolderPath = targetPath + className + "\\";
			File entityTargetFolder = new File(targetFolderPath);
			if (!entityTargetFolder.exists())
				entityTargetFolder.mkdirs();
			if (entityTargetFolder.listFiles().length != 0) {
				for (int i = 0; i < entityTargetFolder.listFiles().length; i++) {
					entityTargetFolder.listFiles()[i].delete();
				}
			}
			List<Field> fields = getFields(datas);
			EntityCreate.create(targetFolderPath, fields);
			ControllerCreate.create(targetFolderPath, fields);
			ServiceCreate.create(targetFolderPath, fields);
			ServiceImplCreate.create(targetFolderPath, fields);
			DaoCreate.create(targetFolderPath, fields);
			MapperCreate.create(targetFolderPath, fields);
			DatabaseCreate.create(tableName, fields);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private static void parseTableName(Map<String, List<List<Object>>> datas) {
		tableName = ExcelPOIUtils.objectToSting(datas.get("Sheet1").get(0).get(1));
	}
	
	private static void parseClassFullName(Map<String, List<List<Object>>> datas) {
		classFullName = ExcelPOIUtils.objectToSting(datas.get("Sheet1").get(0).get(3));
	}
	
	private static void parseObjectName() {
		objectName = StringUtils.dbFieldNameToEntityFieldName(tableName);
	}
	
	public static void parseClassName() {
		String h = objectName.substring(0, 1).toUpperCase();
		className =  h + objectName.substring(1);
	}
	
	private static List<Field> getFields(Map<String, List<List<Object>>> datas) {
		List<Field> fields = new ArrayList<>();
		List<List<Object>> listArray = datas.get("Sheet1");
		for (int i = 2; i < listArray.size(); i++) {
			Field field = new Field();
			if (i == 2)
				field.setPk(true);
			field.setName(ExcelPOIUtils.objectToSting(listArray.get(i).get(0)));
			field.setType(ExcelPOIUtils.objectToSting(listArray.get(i).get(1)));
			field.setDesc(ExcelPOIUtils.objectToSting(listArray.get(i).get(2)));
			field.setDefaultValue(ExcelPOIUtils.objectToSting(listArray.get(i).get(3)));
			fields.add(field);
		}
		return fields;
	}

}
