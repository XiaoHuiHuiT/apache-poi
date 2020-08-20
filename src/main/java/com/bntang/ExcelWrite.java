package com.bntang;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;

import java.io.FileOutputStream;

/**
 * @author ：tangyihao
 * @version ：V1.1.1
 * @program ：BNTang666-Poi
 * @date ：Created in 2020/8/20 15:06
 * @description ：
 */
public class ExcelWrite {

    private static String PATH = "D:\\Devlop\\IDEAProject\\apache-poi-excel\\";

    @Test
    public void Excel03() throws Exception {
        // 1.创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 2.创建一个工作表
        Sheet sheet = workbook.createSheet("03Excel");

        // 3.创建一行
        Row row = sheet.createRow(0);

        // 4.创建一列
        Cell cell = row.createCell(0);
        cell.setCellValue("BNTang666");

        // 5.生成一张表
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "03.xls");
        workbook.write(fileOutputStream);

        // 6.释放资源
        fileOutputStream.flush();
        fileOutputStream.close();
    }

}