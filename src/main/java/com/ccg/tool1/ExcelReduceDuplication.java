package com.ccg.tool1;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelReduceDuplication {

	//测试方法
	public static void main(String[] args) {
		XSSFSheet sheet = getSheet("data");
		List<List<String>> data = readExcelData(sheet);
		if (data == null || data.size() == 0) {
			return;
		}
		Map<Map<String, String>, String> map = new HashMap<>();
		for (List<String> list : data) {
			Map<String, String> m = new HashMap<>();
			for (int i = 0; i < list.size(); i++) {
				if (i > 0) {
					m.put("col" + i, list.get(i));
				}
			}
			map.putIfAbsent(m, list.get(0));
		}

		List<String> pks = data.stream().map(l -> l.get(0)).collect(Collectors.toList());
		Collection<String> values = map.values();
		pks.removeAll(values);
		System.out.println(pks);
	}

	@Test
	public void test() {
		XSSFSheet sheet = getSheet("username");
		readExcelData(sheet);
		//获取第二行第4列
		String cell2 = getExcelDateByIndex(sheet, 1, 3);
		//根据第3列值为“customer23”的这一行，来获取该行第2列的值
		String cell3 = getCellByCaseName(sheet, "customer23", 2, 1);
		System.out.println(cell2);
		System.out.println(cell3);
	}

	private static XSSFSheet getSheet(String sheetName) {
		FileInputStream fileInputStream = null;
		XSSFSheet sheet = null;
		try {
			fileInputStream = new FileInputStream("src/main/java/com/ccg/tool1/data.xlsx");
			XSSFWorkbook sheets = new XSSFWorkbook(fileInputStream);
			//获取sheet
			sheet = sheets.getSheet(sheetName);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return sheet;
	}

	/**
	 * 根据行和列的索引获取单元格的数据
	 *
	 * @param row
	 * @param column
	 * @return
	 */
	public static String getExcelDateByIndex(XSSFSheet sheet, int row, int column) {
		XSSFRow row1 = sheet.getRow(row);
		String cell = row1.getCell(column).toString();
		return cell;
	}

	/**
	 * 根据某一列值为“******”的这一行，来获取该行第x列的值
	 *
	 * @param caseName
	 * @param currentColumn 当前单元格列的索引
	 * @param targetColumn  目标单元格列的索引
	 * @return
	 */
	public static String getCellByCaseName(XSSFSheet sheet, String caseName, int currentColumn, int targetColumn) {
		String operateSteps = "";
		//获取行数
		int rows = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < rows; i++) {
			XSSFRow row = sheet.getRow(i);
			String cell = row.getCell(currentColumn).toString();
			if (cell.equals(caseName)) {
				operateSteps = row.getCell(targetColumn).toString();
				break;
			}
		}
		return operateSteps;
	}

	//打印excel数据
	public static List<List<String>> readExcelData(XSSFSheet sheet) {
		List<List<String>> data = new ArrayList<>();
		//获取行数
		int rowNum = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < rowNum; i++) {
			ArrayList<String> list = new ArrayList<>();
			//获取列数
			XSSFRow row = sheet.getRow(i);
			int columnNum = row.getPhysicalNumberOfCells();
			for (int j = 0; j < columnNum; j++) {
				String cell = row.getCell(j).toString();
				list.add(cell);
				System.out.print(cell + (j < columnNum - 1 ? ", " : "\n"));
			}
			data.add(list);
		}
		return data;
	}

}
