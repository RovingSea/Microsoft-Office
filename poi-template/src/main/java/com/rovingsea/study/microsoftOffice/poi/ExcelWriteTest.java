package com.rovingsea.study.microsoftOffice.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author Haixin Wu
 * @since 1.0
 */
public class ExcelWriteTest {

    public static String PATH = "./src/main/resources/";

    @Test
    public void writeTest03() throws IOException {
        // 1、创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        // 2、创建一个工作表
        Sheet sheet = workbook.createSheet("1-sheet");
        // 3、创建一个行
        Row row1 = sheet.createRow(0);
        // 4、创建一个单元格
        Cell cell11 = row1.createCell(0);
        // 5、单元格中的数据
        cell11.setCellValue("(1,1)");
        // 6、同理：第一行第二列
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("(1,2)");
        // 7、同理：第二行第一列
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("(2,1)");
        // 8、同理：第二行第二列
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));
        // 9、生成xls
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "03.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("生成完毕");
    }

    @Test
    public void writeTest07() throws IOException {
        // 1、创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        // 2、创建一个工作表
        Sheet sheet = workbook.createSheet("1-sheet");
        // 3、创建一个行
        Row row1 = sheet.createRow(0);
        // 4、创建一个单元格
        Cell cell11 = row1.createCell(0);
        // 5、单元格中的数据
        cell11.setCellValue("(1,1)");
        // 6、同理：第一行第二列
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("(1,2)");
        // 7、同理：第二行第一列
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("(2,1)");
        // 8、同理：第二行第二列
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));
        // 9、生成xlsx
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("生成完毕");
    }

    /**
     * 较快，但是超过65536条后会抛出异常
     */
    @Test
    public void writeBigDataTest03() throws IOException {
        long begin = System.currentTimeMillis();

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int colNum = 0; colNum < 10; colNum++) {
                Cell cell = row.createCell(colNum);
                cell.setCellValue(colNum);
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "bigdata03.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("生成完毕");

        long end = System.currentTimeMillis();
        System.out.println("03版本消耗时间：" + (end - begin));
    }

    /**
     * 较慢，但是没有行列限制 <br>
     * 可以使用SXSSFWorkbook （Super XSSFWorkbook）解决速度慢的问题 <br>
     * 但是会生成临时文件，可以使用{@link SXSSFWorkbook#dispose()}清除临时文件
     *
     */
    @Test
    public void writeBigDataTest07() throws IOException {
        long begin = System.currentTimeMillis();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int colNum = 0; colNum < 10; colNum++) {
                Cell cell = row.createCell(colNum);
                cell.setCellValue(colNum);
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "bigdata07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("生成完毕");

        long end = System.currentTimeMillis();
        System.out.println("07版本消耗时间：" + (end - begin));
    }
}

