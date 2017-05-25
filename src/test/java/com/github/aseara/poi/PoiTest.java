package com.github.aseara.poi;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import static org.apache.poi.ss.usermodel.CellType.STRING;

/**
 * Created by qiujingde on 2017/2/21.
 * POI简单测试
 */
public class PoiTest {

    private static final Logger LOG = LoggerFactory.getLogger(PoiTest.class);

    @Test
    public void read() throws IOException {
        InputStream xlsInput = getClass().getResourceAsStream("/test.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(xlsInput);
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(1);
        XSSFCell cell = row.getCell(1);
        LOG.info(cell.getStringCellValue());
    }

    @Test
    public void write() throws IOException {
        String outFile = "target/out/test.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet();
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);
        cell.setCellType(STRING);
        cell.setCellValue("OK");

        try (FileOutputStream outStream = new FileOutputStream(outFile)) {
            workbook.write(outStream);
        }
    }

}
