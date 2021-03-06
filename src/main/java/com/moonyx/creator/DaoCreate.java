package com.moonyx.creator;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

public class DaoCreate {

	public static void create(String targetFolderPath, List<Field> fields) {
		String packageName = Creator.classFullName.replace("entity." + Creator.className, "dao");
		
		File daoModel = new File(Creator.modelPath + "dao.java");
		File daoFile = new File(targetFolderPath + Creator.className + "Dao.java");
		BufferedReader reader = null;
		FileWriter fw = null;
		BufferedWriter writer = null;
		try {
			reader = new BufferedReader(new FileReader(daoModel));
			daoFile.createNewFile();
			fw = new FileWriter(daoFile);
			writer = new BufferedWriter(fw);
			String tempString = null;
			while ((tempString = reader.readLine()) != null) {
				String s = tempString.replaceAll("<PackageName>", packageName);
				s = s.replaceAll("<EntityClassFullName>", Creator.classFullName);
				s = s.replaceAll("<EntityClassName>", Creator.className);
				s = s.replaceAll("<EntityObjectName>", Creator.objectName);
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

}
