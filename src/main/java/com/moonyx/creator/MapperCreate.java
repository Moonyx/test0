package com.moonyx.creator;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

import com.moonyx.utils.StringUtils;

public class MapperCreate {
	
	public static void create(String targetFolderPath, List<Field> fields) {
		String mapperName = Creator.classFullName.replace("entity", "dao");
		
		File mapperModel = new File(Creator.modelPath + "mapper.xml");
		File mapperFile = new File(targetFolderPath + Creator.className + "Mapper.xml");
		BufferedReader reader = null;
		FileWriter fw = null;
		BufferedWriter writer = null;
		try {
			reader = new BufferedReader(new FileReader(mapperModel));
			mapperFile.createNewFile();
			fw = new FileWriter(mapperFile);
			writer = new BufferedWriter(fw);
			String tempString = null;
			while ((tempString = reader.readLine()) != null) {
				String s = tempString.replaceAll("<EntityClassName>", Creator.className);
				s = s.replaceAll("<EntityObjectName>", Creator.objectName);
				s = s.replaceAll("<MapperName>", mapperName);
				s = s.replaceAll("<TableName>", Creator.tableName);
				s = s.replaceAll("<EntityClassFullName>", Creator.classFullName);
				if (s.indexOf("<ResultMapIdField>") > -1)
					s = s.replace("<ResultMapIdField>", getResultMapId(fields));
				if (s.indexOf("<ResultMapFields>") > -1)
					s = s.replace("<ResultMapFields>", getResultMap(fields, writer));
				if (s.indexOf("<BaseColumnList>") > -1)
					s = s.replace("<BaseColumnList>", getBaseColumnList(fields));
				if (s.indexOf("<WhereIdSql>") > -1)
					s = s.replace("<WhereIdSql>", getWhereIdSql(fields));
				if (s.indexOf("<InsertSql>") > -1)
					s = s.replace("<InsertSql>", getInsertSql(fields));
				if (s.indexOf("<UpdateSql>") > -1)
					s = s.replace("<UpdateSql>", getUpdateSql(fields));
				if (s.indexOf("<DeleteSql>") > -1)
					s = s.replace("<DeleteSql>", getDeleteSql(fields));
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
	
	private static String getResultMapId(List<Field> fields) {
		String column = fields.get(0).getName();
		String property = StringUtils.dbFieldNameToEntityFieldName(column);
		String type = fields.get(0).getType();
		int index = type.indexOf("(");
		if (index > -1)
			type = type.substring(0, index);
		return "<id column=\"" + column + "\" property=\"" + property + "\" jdbcType=\"" + parseJdbcType(type) + "\" />";
	}
	
	private static String getResultMap(List<Field> fields, BufferedWriter writer) throws IOException {
		for (int i = 1; i < fields.size(); i++) {
			String column = fields.get(i).getName();
			String property = StringUtils.dbFieldNameToEntityFieldName(column);
			String type = fields.get(i).getType();
			int index = type.indexOf("(");
			if (index > -1)
				type = type.substring(0, index);
			writer.write("    <result column=\"" + column + "\" property=\"" + property + "\" jdbcType=\"" + parseJdbcType(type) + "\" />");
			writer.newLine();
		}
		return "";
	}
	
	private static String parseJdbcType(String dbType) {
		dbType = dbType.toUpperCase();
		switch (dbType) {
			case "INT" : return "INTEGER";
			case "LONG" : return "INTEGER";
			case "DATETIME" : return "TIMESTAMP";
			default : break;
		}
		return dbType;
	}
	
	private static String getBaseColumnList(List<Field> fields) {
		StringBuffer s = new StringBuffer();
		for (int i = 0; i < fields.size(); i++) {
			if (i > 0)
				s.append(", ");
			s.append(fields.get(i).getName());
		}
		return s.toString();
	}
	
	private static String getInsertSql(List<Field> fields) {
//		INSERT INTO tableName (author, title,url) VALUES (#{author}, #{title}, #{url})
		StringBuffer s = new StringBuffer();
		s.append("INSERT INTO " + Creator.tableName + " (");
		for (int i = 0; i < fields.size(); i++) {
			if (i > 0)
				s.append(", ");
			s.append(fields.get(i).getName());
		}
		s.append(") VALUES (");
		for (int i = 0; i < fields.size(); i++) {
			if (i > 0)
				s.append(", ");
			s.append("#{" + StringUtils.dbFieldNameToEntityFieldName(fields.get(i).getName()) + "}");
		}
		s.append(")");
		return s.toString();
	}
	
	private static String getUpdateSql(List<Field> fields) {
//		UPDATE tableName SET author = #{author},title = #{title},url = #{url} WHERE id = #{id}
		StringBuffer s = new StringBuffer();
		s.append("UPDATE " + Creator.tableName + " SET ");
		for (int i = 0; i < fields.size(); i++) {
			if (i > 0)
				s.append(", ");
			s.append(fields.get(i).getName());
			s.append(" = #{" + StringUtils.dbFieldNameToEntityFieldName(fields.get(i).getName()) + "}");
		}
		s.append(getWhereIdSql(fields));
		return s.toString();
	}
	
	private static String getWhereIdSql(List<Field> fields) {
		return " WHERE " + fields.get(0).getName() + " = #{" + StringUtils.dbFieldNameToEntityFieldName(fields.get(0).getName()) + "}";
	}
	
	private static String getDeleteSql(List<Field> fields) {
//		DELETE FROM TableName WHERE id in
		return "DELETE FROM " + Creator.tableName + " WHERE " + fields.get(0).getName() + " in ";
	}

}
