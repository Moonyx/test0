package com.moonyx.creator;

import java.util.List;

public class DatabaseCreate {
	
	public static void create(String tableName, List<Field> fields) {
		System.out.println("create table " + tableName + " (");
		for (int i = 0; i < fields.size(); i++) {
			Field field = fields.get(i);
			System.out.println("" + field.getName() + " " + field.getType() + " COLLATE utf8_bin DEFAULT NULL COMMENT '" + field.getDesc() + "',");
		}
		System.out.println("PRIMARY KEY (" + fields.get(0).getName() + ")");
		System.out.println(") ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_bin;");
	}

}
