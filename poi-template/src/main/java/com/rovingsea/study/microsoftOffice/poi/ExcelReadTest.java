package com.rovingsea.study.microsoftOffice.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * @author Haixin Wu
 * @since 1.0
 */
public class ExcelReadTest {

    public static String PATH = "./src/main/resources/";

    @Test
    public void readTest03() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "03.xls");
        // 1、从输入流得到工作簿
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        // 2、得到表
        Sheet sheet = workbook.getSheetAt(0);
        // 3、得到行
        Row row = sheet.getRow(0);
        // 4、得到列
        Cell cell = row.getCell(0);
        // 5、读取值的时候要注意类型，该工具类服务自适应匹配，可惜~
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();
    }

    @Test
    public void readTest07() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "07.xlsx");
        // 1、从输入流得到工作簿
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        // 2、得到表
        Sheet sheet = workbook.getSheetAt(0);
        // 3、得到行
        Row row = sheet.getRow(0);
        // 4、得到列
        Cell cell = row.getCell(0);
        // 5、读取值的时候要注意类型，该工具类服务自适应匹配，可惜~
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();
    }

}

