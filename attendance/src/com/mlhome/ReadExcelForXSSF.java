package com.mlhome;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcelForXSSF {  
    public void read() {  
        File file = new File("C:\\Users\\malei\\Desktop\\yyb.xlsx");  
        InputStream inputStream = null;  
        Workbook workbook = null;  
        try {  
            inputStream = new FileInputStream(file);  
            workbook = WorkbookFactory.create(inputStream);  
            inputStream.close();  
            //���������  
            Sheet sheet = workbook.getSheetAt(0);  
            //������  
            int rowLength = sheet.getLastRowNum()+1;  
            //���������  
            Row row = sheet.getRow(0);  
            //������  
            int colLength = row.getLastCellNum();  
            //�õ�ָ���ĵ�Ԫ��  
            Cell cell = row.getCell(0);  
            //�õ���Ԫ����ʽ  
            CellStyle cellStyle = cell.getCellStyle();  
            System.out.println("������" + rowLength + ",������" + colLength);  
            for (int i = 0; i < rowLength; i++) {  
                row = sheet.getRow(i);  
                for (int j = 0; j < colLength; j++) {  
                    cell = row.getCell(j);  
                    //Excel����Cell�в�ͬ�����ͣ���������ͼ��һ���������͵�Cell��ȡ��һ���ַ���ʱ���п��ܱ��쳣��  
                    //Cannot get a STRING value from a NUMERIC cell  
                    //�����е���Ҫ����Cell�������ΪString��ʽ  
                    if (cell != null)  
                        cell.setCellType(CellType.STRING);  
  
                    //��Excel�����޸�  
                    if (i > 0 && j == 1)  
                        cell.setCellValue("1000");  
                    System.out.print(cell.getStringCellValue() + "\t");  
                }  
                System.out.println();  
            }  
            //���޸ĺõ����ݱ���  
            OutputStream out = new FileOutputStream(file);  
            workbook.write(out);  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    }  
  
    public static void main(String[] args) {  
        new ReadExcelForXSSF().read();  
    }  
}  