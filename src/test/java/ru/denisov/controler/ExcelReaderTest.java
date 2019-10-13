package ru.denisov.controler;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import ru.denisov.model.Lesson;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.NoSuchFileException;
import java.util.Properties;

import static org.junit.jupiter.api.Assertions.*;

class ExcelReaderTest {
    private static Properties prop = new Properties();

    @BeforeAll
    static void initProps() throws IOException {
        prop.load(new FileInputStream("src/main/resources/kompas.properties"));
    }

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
        showOddWeek(sheet, 9);
    }

    void showOddWeek(XSSFSheet sheet, int cell){
        for(int i = 0; i < Integer.parseInt(String.valueOf(prop.get("DAYS"))); i++){
            System.out.println("День " + (i + 1));
            for(int j = 0; j < Integer.parseInt(String.valueOf(prop.get("LESSONS"))); j++){
                System.out.println("Пара " + (j + 1) + ": "
                        + getLesson(sheet,Integer.parseInt(String.valueOf(prop.get("ROW_PADDING")))
                        + Integer.parseInt(String.valueOf(prop.get("DAY_SHIFT"))) * i
                        + Integer.parseInt(String.valueOf(prop.get("LESSON_SHIFT"))) * j, cell));
            }
        }
    }

    String getLesson(XSSFSheet sheet, int row, int cell){
        XSSFRow row_t = sheet.getRow(row);
        Lesson lesson = new Lesson();
        lesson.setSubject(row_t.getCell(cell).getStringCellValue());
        lesson.setType(row_t.getCell(cell + 1).getStringCellValue());
        lesson.setTeacher(row_t.getCell(cell + 2).getStringCellValue());
        lesson.setAuditory(row_t.getCell(cell + 3).getStringCellValue());
//        System.out.println(lesson.toString());
//        String s_lesson = row_t.getCell(cell).getStringCellValue() + " " + row_t.getCell(cell + 1).getStringCellValue();
//        if (s_lesson.equals(" ")){
//            s_lesson = "-";
//        }
        return lesson.getSubject().equals("") ? "-" : lesson.toString();
    }
}