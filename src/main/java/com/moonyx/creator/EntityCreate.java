package com.moonyx.creator;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

import com.moonyx.utils.StringUtils;

public class EntityCreate {

	public static void create(String targetFolderPath, List<Field> fields) {
		String packageName = Creator.classFullName.replace("." + Creator.className, "");
		
		File entityModel = new File(Creator.modelPath + "entity.java");
		File entityFile = new File(targetFolderPath + Creator.className + ".java");
		BufferedReader reader = null;
		FileWriter fw = null;
		BufferedWriter writer = null;
		try {
			reader = new BufferedReader(new FileReader(entityModel));
			entityFile.createNewFile();
			fw = new FileWriter(entityFile);
			writer = new BufferedWriter(fw);
			String tempString = null;
			while ((tempString = reader.readLine()) != null) {
				String s = tempString.replaceAll("<PackageName>", packageName);
				s = s.replaceAll("<EntityClassName>", Creator.className);
				s = s.replaceAll("<EntityObjectName>", Creator.objectName);
				if (s.indexOf("<entityBody>") > -1) {
					s = s.replaceAll("<entityBody>", "");
					writer.newLine();
					for (int i = 0; i < fields.size(); i++) {
						Field field = fields.get(i);
						String fieldDeclaration = "\tprivate " + getFieldType(field.getType()) + " " + StringUtils.dbFieldNameToEntityFieldName(field.getName()) + ";";
						writer.write(fieldDeclaration);
						writer.newLine();
						writer.newLine();
					}
					for (int i = 0; i < fields.size(); i++) {
						Field field = fields.get(i);
						String getFunction1 = "\tpublic " + getFieldType(field.getType()) + " get" + parseFieldName(field) + "() {";
						writer.write(getFunction1);
						writer.newLine();
						String getFunction2 = "\t\treturn " + StringUtils.dbFieldNameToEntityFieldName(field.getName())+ ";";
						writer.write(getFunction2);
						writer.newLine();
						String getFunction3 = "\t}";
						writer.write(getFunction3);
						writer.newLine();
						writer.newLine();

						String setFunction1 = "\tpublic void set" + parseFieldName(field) + "(" + getFieldType(field.getType()) + " "
								+ StringUtils.dbFieldNameToEntityFieldName(field.getName()) + ") {";
						writer.write(setFunction1);
						writer.newLine();
						String setFunction2 = "\t\tthis." + StringUtils.dbFieldNameToEntityFieldName(field.getName())
								+ " = " + StringUtils.dbFieldNameToEntityFieldName(field.getName()) + ";";
						writer.write(setFunction2);
						writer.newLine();
						String setFunction3 = "\t}";
						writer.write(setFunction3);
						writer.newLine();
						if (i < fields.size() - 1)
							writer.newLine();
//						public String getLtCustRequestId() {
//							return ltCustRequestId;
//						}
			//
//						public void setLtCustRequestId(String ltCustRequestId) {
//							this.ltCustRequestId = ltCustRequestId;
//						}
					}
				}
				writer.write(s);
				writer.newLine();
			}
			reader.close();
			writer.flush();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (reader != null) {
				try {
					reader.close();
				} catch (IOException e1) {
				}
			}
			if (fw != null) {
				try {
					fw.close();
				} catch (IOException e) {
				}
			}
			if (writer != null) {
				try {
					writer.close();
				} catch (IOException e) {
				}
			}
		}
	}
	
	private static String getFieldType(String dbFieldType) {
		dbFieldType = dbFieldType.toUpperCase();
		if (dbFieldType.indexOf("VARCHAR") > -1)
			return "String";
		if (dbFieldType.indexOf("DATE") > -1)
			return "Date";
		if (dbFieldType.indexOf("INT") > -1)
			return "int";
		if (dbFieldType.indexOf("FLOAT") > -1)
			return "float";
		if (dbFieldType.indexOf("DOUBLE") > -1)
			return "double";
		return "NULL";
	}
	
	public static String parseFieldName(Field field) {
		String fieldName = StringUtils.dbFieldNameToEntityFieldName(field.getName());
		String h = fieldName.substring(0, 1).toUpperCase();
		return h + fieldName.substring(1);
	}

}
