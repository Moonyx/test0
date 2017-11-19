package com.moonyx.utils;

public class StringUtils {
	
	public static String dbFieldNameToEntityFieldName(String text) {
		if (text == null || "".equals(text))
			return text;
		int beginIndex = 0;
		int _index = text.indexOf("_");
		if (_index == -1)
			return text.toLowerCase();
		String targetText = "";
		while (_index > -1) {
			String s = text.substring(beginIndex, _index);
			targetText += s.toLowerCase();
			String h = text.substring(_index + 1, _index + 2);
			targetText += h.toUpperCase();
			beginIndex = _index + 2;
			_index = text.indexOf("_", _index + 2);
			if (_index == -1) {
				targetText += text.substring(beginIndex, text.length()).toLowerCase();
			}
		}
		return targetText;
    }
	
	public static String entityFieldNameToDbFieldName(String text) {
		if (text == null || "".equals(text))
			return text;
		String targetText = "";
		for (int i = 0; i < text.length(); i++) {
			char c = text.charAt(i);
			if (Character.isUpperCase(c))
				targetText += "_";
			targetText += c;
		}
		return targetText.toUpperCase();
    }

}
