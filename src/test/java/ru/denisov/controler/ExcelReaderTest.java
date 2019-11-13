package ru.denisov.controler;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import ru.denisov.model.Lesson;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.NoSuchFileException;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

class ExcelReaderTest {
    private static Properties prop = new Properties();
    public static LinkedList<Lesson> lessons = new LinkedList<>();
    public static HashSet<String> auditories = new HashSet<>();
    public static HashSet<String> teachers = new HashSet<>();

    @BeforeAll
    static void initProps() throws IOException {
        prop.load(new FileInputStream("src/main/resources/kompas.properties"));
    }

    @Test
    void getAllLessons() throws NoSuchFileException {
        for(File excel: Objects.requireNonNull(new File("src/main/resources/lessons/").listFiles())){
            XSSFWorkbook wb = ExcelReader.readWorkbook(excel);
            XSSFSheet sheet = wb.getSheetAt(0);
            parseSheet(sheet);
        }
        auditoryTime();
    }

    static void parseSheet(XSSFSheet sheet){
        for(int group_index = 0; !getGroupName(sheet, group_index).isEmpty(); group_index++){
            for(int i = 0; i < Integer.parseInt(String.valueOf(prop.get("DAYS"))); i++){
                for(int j = 0; j < Integer.parseInt(String.valueOf(prop.get("LESSONS"))); j++) {
                    getLesson(sheet,Integer.parseInt(String.valueOf(prop.get("ROW_PADDING")))
                            + Integer.parseInt(String.valueOf(prop.get("DAY_SHIFT"))) * i
                            + Integer.parseInt(String.valueOf(prop.get("LESSON_SHIFT"))) * j, getGroupShift(group_index), j, group_index);
                    getLesson(sheet,Integer.parseInt(String.valueOf(prop.get("ROW_PADDING")))
                            + Integer.parseInt(String.valueOf(prop.get("DAY_SHIFT"))) * i
                            + Integer.parseInt(String.valueOf(prop.get("LESSON_SHIFT"))) * j + 1, getGroupShift(group_index), j, group_index);
                }
            }
        }
    }

    @Test
    void auditoryTime(){
        String audit = "А-11";
        HashSet<Lesson> lesson_in_a = new HashSet<>();
        for(Lesson lesson_t: lessons){
            if (audit.equals(lesson_t.getAuditory())){
                lesson_in_a.add(lesson_t);
            }
        }
        System.out.println(lesson_in_a);
    }


    @Test
    void readWorkbook() throws NoSuchFileException {
        File excel = new File("src/main/resources/lessons/IIT_1k_19_20_osen.xlsx");
        assertTrue(excel.isFile());
        XSSFWorkbook wb = ExcelReader.readWorkbook(excel);
        assertNotNull(wb);
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row = sheet.getRow(11);
        String lesson = row.getCell(5).getStringCellValue();
//        assertEquals("БЖД", lesson);
        for (int i = 0; !getGroupName(sheet, i).isEmpty(); i++) {
            System.out.println(getGroupName(sheet, i));
            showOddWeek(sheet, i);
            showEvenWeek(sheet, i);
        }
    }

    @Test
    void groupName() throws NoSuchFileException {
        File excel = new File("src/main/resources/lessons/IIT_3k_19_20_osen.xlsx");
        XSSFWorkbook wb = ExcelReader.readWorkbook(excel);
        XSSFSheet sheet = wb.getSheetAt(0);
        assertEquals("ИВБО-01-17", getGroupName(sheet, 3));
    }

    public static String getGroupName(XSSFSheet sheet, int group_index){
        XSSFRow row = sheet.getRow(1);
        return getCellString(row.getCell(getGroupShift(group_index)));
    }

    void showOddWeek(XSSFSheet sheet, int group_index){
        System.out.println("Нечетная неделя");
        for(int i = 0; i < Integer.parseInt(String.valueOf(prop.get("DAYS"))); i++){
            System.out.println("День " + (i + 1));
            for(int j = 0; j < Integer.parseInt(String.valueOf(prop.get("LESSONS"))); j++){
                System.out.println("Пара " + (j + 1) + ": "
                        + getLesson(sheet,Integer.parseInt(String.valueOf(prop.get("ROW_PADDING")))
                        + Integer.parseInt(String.valueOf(prop.get("DAY_SHIFT"))) * i
                        + Integer.parseInt(String.valueOf(prop.get("LESSON_SHIFT"))) * j, getGroupShift(group_index), j, group_index));
            }
        }
    }

    void showEvenWeek(XSSFSheet sheet, int group_index){
        System.out.println("Четная неделя");
        for(int i = 0; i < Integer.parseInt(String.valueOf(prop.get("DAYS"))); i++){
            System.out.println("День " + (i + 1));
            for(int j = 0; j < Integer.parseInt(String.valueOf(prop.get("LESSONS"))); j++){
                System.out.println("Пара " + (j + 1) + ": "
                        + getLesson(sheet,Integer.parseInt(String.valueOf(prop.get("ROW_PADDING")))
                            + Integer.parseInt(String.valueOf(prop.get("DAY_SHIFT"))) * i
                            + Integer.parseInt(String.valueOf(prop.get("LESSON_SHIFT"))) * j + 1, getGroupShift(group_index), j, group_index));
            }
        }
    }

    @Test
    void nullCell() throws NoSuchFileException {
        File excel = new File("src/main/resources/IIT_1k_19_20_osen.xlsx");
        XSSFWorkbook wb = ExcelReader.readWorkbook(excel);
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row_t = sheet.getRow(43);
        System.out.println("'"+getCellString(row_t.getCell(24))+"'");
    }

    static String getCellString(XSSFCell cell){
        try { return cell.getStringCellValue(); }
        catch (NullPointerException e){ return ""; }
    }

    public static String getLesson(XSSFSheet sheet, int row, int cell, int lesson_number, int group_index){
        XSSFRow row_t = sheet.getRow(row);
        Lesson lesson = new Lesson();
        try {
            lesson.setSubject(getCellString(row_t.getCell(cell)));
            lesson.setType(getCellString(row_t.getCell(cell + 1)));
            lesson.setTeacher(getCellString(row_t.getCell(cell + 2)));
            teachers.add(lesson.getTeacher());
            lesson.setAuditory(getCellString(row_t.getCell(cell + 3)));
            auditories.add(lesson.getAuditory());
            lesson.setLessonNumber(lesson_number);
        }
        catch (IllegalStateException e){
            System.out.println("Cannot read cells at " + getGroupName(sheet, group_index) + " " + row + " " + cell);
            e.printStackTrace();
            throw e;
        }
        if (!lesson.getSubject().equals("")){
            lessons.add(lesson);
            return lesson.toString();
        }
        else return "-";
//        System.out.println(lesson.toString());
//        String s_lesson = row_t.getCell(cell).getStringCellValue() + " " + row_t.getCell(cell + 1).getStringCellValue();
//        if (s_lesson.equals(" ")){
//            s_lesson = "-";
//        }
//        return lesson.getSubject().equals("") ? "-" : lesson.toString();
    }

    public static int getGroupShift(int group_index){
        int result = Integer.parseInt(prop.get("GROUP_PADDING").toString());
        for(int i = 0; i < group_index; i++){
            result += Integer.parseInt(prop.get("GROUP_SHIFT").toString().split(",")[i % 3]);
        }
        return result;
    }
}