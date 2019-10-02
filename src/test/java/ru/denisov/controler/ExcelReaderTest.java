package ru.denisov.controler;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.nio.file.NoSuchFileException;

import static org.junit.jupiter.api.Assertions.*;

class ExcelReaderTest {

    @org.junit.jupiter.api.Test
    void readWorkbook() throws NoSuchFileException {
        File excel = new File("src/main/resources/IIT_3k_19_20_osen.xlsx");
        assertTrue(excel.isFile());
        System.out.println(excel.getAbsoluteFile());
        XSSFWorkbook wb = ExcelReader.readWorkbook(excel);
        assertNotNull(wb);
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row = sheet.getRow(11);
        String lesson = row.getCell(5).getStringCellValue();
        System.out.println(lesson);
        assertEquals("БЖД", lesson);
    }
}