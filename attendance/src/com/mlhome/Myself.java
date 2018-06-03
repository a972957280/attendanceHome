package com.mlhome;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.model.WorkbookRecordList;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Myself {
	public void read() {
		//读取文件
		File file = new File("C:\\\\Users\\\\malei\\\\Desktop\\\\yyb.xlsx");
		//输入流
		InputStream put =null;
		Workbook wb=null;
		try {
			put = new FileInputStream(file);
			wb=WorkbookFactory.create(put);
			//工作表对象
			Sheet sheet=wb.getSheetAt(0);
			//总行数
			int rowLength= sheet.getLastRowNum()+1;
			//工作表的列
			Row row=sheet.getRow(0);
			//总列数
			int colLength = row.getLastCellNum();
			//得到指定的单元格
			Cell cell= row.getCell(0);
			System.out.println("行："+rowLength+",列："+colLength);
			for (int i = 0; i < rowLength; i++) {
				row=sheet.getRow(i);
				for (int j = 0; j < colLength; j++) {
					cell=row.getCell(j);
//					System.out.println(cell.getStringCellValue()+"\t");
					while (j==7) {
						System.out.println(cell.getStringCellValue());
						break;
					}
				}
			}
		} catch (Exception e) {
			e.getMessage();
		}
		
	}
	public static void main(String[] args) {
		new Myself().read();
	}
}
