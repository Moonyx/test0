package com.moonyx.creator;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import com.moonyx.utils.ExcelPOIUtils;
import com.moonyx.utils.StringUtils;

public class Creator {
	
	public static String path = "D:\\work\\create\\claw\\";
	public static String modelPath = "D:\\work\\create\\claw\\model\\";
	public static String targetPath = "D:\\work\\create\\claw\\target\\";
	
	public static SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
	
	public static void main(String[] args) {
		File file = new File(path + "claw.xlsx");
		try {
			Map<String, List<List<Object>>> datas = ExcelPOIUtils.readExcelSheetDatas(new FileInputStream(file), file.getName());
			String tableName = getTableName(datas);
			String entityObjectName = StringUtils.dbFieldNameToEntityFieldName(tableName);
			String entityClassName = getEntityName(entityObjectName);
			String targetFolderPath = targetPath + entityClassName + "\\" + sdf.format(new Date()) + "\\";
			File entityTargetFolder = new File(targetFolderPath);
			if (!entityTargetFolder.exists())
				entityTargetFolder.mkdirs();
			List<Field> fields = getFields(datas);
			EntityCreate.createEntity(targetFolderPath, entityClassName, entityObjectName, fields);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private static String getTableName(Map<String, List<List<Object>>> datas) {
		return ExcelPOIUtils.objectToSting(datas.get("Sheet1").get(0).get(1));
	}
	
	public static String getEntityName(String entityObjectName) {
		String h = entityObjectName.substring(0, 1).toUpperCase();
		return h + entityObjectName.substring(1);
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
