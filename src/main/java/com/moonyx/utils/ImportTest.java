package com.moonyx.utils;

import java.io.File;
import java.text.DecimalFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ImportTest {
	
	public static void main(String[] args) {
		try {
//			String importFilePath = "C:\\Users\\user\\Documents\\Tencent Files\\515230772\\FileRecv\\NS-HK-001-2018.xlsx";
//			String importFilePath = "C:\\Users\\user\\Documents\\Tencent Files\\515230772\\FileRecv\\NS-HK-001-2018r.xlsx";
			String importFilePath = "C:\\Users\\user\\Documents\\Tencent Files\\515230772\\FileRecv\\NS-HK-002-2018r.xlsx";
			Workbook workbook = WorkbookFactory.create(new File(importFilePath));
			Sheet sheet = workbook.getSheetAt(0);
			StringBuffer s = new StringBuffer();
			for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
				Row row = sheet.getRow(rowNum);
				if (row == null) {
					break;
				}
				Cell cell0 = row.getCell(0);
				Cell cell1 = row.getCell(1);
//				Cell cell2 = row.getCell(2);
				Cell cell3 = row.getCell(3);
				Cell cell4 = row.getCell(4);
//				Cell cell5 = row.getCell(5);
//				Cell cell6 = row.getCell(6);
//				Cell cell7 = row.getCell(7);
				Cell cell8 = row.getCell(8);
				Cell cell9 = row.getCell(9);
//				Cell cell10 = row.getCell(10);
				Cell cell11 = row.getCell(11);
				Cell cell12 = row.getCell(12);
				Cell cell13 = row.getCell(13);
				Cell cell14 = row.getCell(14);
				Cell cell15 = row.getCell(15);
				Cell cell16 = row.getCell(16);
				System.out.println(cell0.getStringCellValue());
				s.append("<WBK_PDE_ITEM_ORG>");
				s.append("<radeclno/>");
				s.append("<idx>" + rowNum + "</idx>");
				s.append("<gcattr1>" + cell0.getStringCellValue() + "</gcattr1>");
				s.append("<sub_fld2>" + new DecimalFormat("0").format(cell16.getNumericCellValue()) + "</sub_fld2>");
				s.append("<hs_code>" + new DecimalFormat("0").format(cell11.getNumericCellValue()) + "</hs_code>");
				s.append("<hs_gname>" + cell1.getStringCellValue() + "</hs_gname>");
				s.append("<hs_spec></hs_spec>");
				s.append("<qty>" + new DecimalFormat("0").format(cell3.getNumericCellValue()) + "</qty>");
				s.append("<grossweight>" + new DecimalFormat("0.###").format(cell13.getNumericCellValue()) + "</grossweight>");
				s.append("<weight>" + new DecimalFormat("0.###").format(cell12.getNumericCellValue()) + "</weight>");
				s.append("<volume>" + new DecimalFormat("0.###").format(cell14.getNumericCellValue()) + "</volume>");
				s.append("<hs_qty>" + new DecimalFormat("0").format(cell4.getNumericCellValue()) + "</hs_qty>");
				s.append("<app_unit></app_unit>");
				s.append("<hs_lawqty1></hs_lawqty1>");
				s.append("<hs_lawunit1></hs_lawunit1>");
				s.append("<hs_lawqty2/>");
				s.append("<hs_lawunit2/>");
				s.append("<purfcode>502</purfcode>");
				s.append("<purupric>" + new DecimalFormat("0.##").format(cell8.getNumericCellValue()) + "</purupric>");
				s.append("<fcy>" + new DecimalFormat("0.##").format(cell9.getNumericCellValue()) + "</fcy>");
				s.append("<tax_mode>1</tax_mode>");
				s.append("<hs_origin>142</hs_origin>");
				s.append("<hs_destination>116</hs_destination>");
				s.append("<hs_lawcheck>0</hs_lawcheck>");
				s.append("<impcustomno>" + cell15.getStringCellValue() + "</impcustomno>");
				s.append("</WBK_PDE_ITEM_ORG>");
			}
			System.out.println(s.toString());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
