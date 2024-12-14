package com.excelparser;

import com.aspose.cells.*;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.ArrayList;
import java.util.List;

/**
 * Класс ExcelParser создан для обработки данных их Excel-файла,
 * анализа информации о студентах и их оценках, создания сводных результатов
 * и построения диаграммы
 */

public class ExcelParser {

    private static final Logger logger = LogManager.getLogger(ExcelParser.class);

    /**
     * Запускает обработку данных из файла
     */

    public static void start(String inputFilePath, String outputFilePath) {
        logger.info("Запуск приложения");

        //String inputFilePath = "src/main/resources/input.xlsx";
        //String outputFilePath = "src/main/resources/output.xlsx";

        try {
            ExcelParser app = new ExcelParser();
            app.processExcel(inputFilePath, outputFilePath);

            logger.info("Готово");
        } catch (Exception e) {
            logger.error("Ошибка выполнения приложения");
        }
    }

    /**
     * Обрабатывает данные из Excel-файла, сохраняет результаты
     *
     * @param inputFilePath  Путь к исходному Excel-файлу
     * @param outputFilePath Путь к конечному Excel-файлу
     * @throws Exception В случае ошибки при обработки данных
     */
    private void processExcel(String inputFilePath, String outputFilePath) throws Exception {
        // Загрузка книги Excel
        Workbook workbook = new Workbook(inputFilePath);

        // Получение листа
        Worksheet sheet = workbook.getWorksheets().get(0);

        // Списки для хранения данных о студентах с разными оценками
        List<String> excellentStudents = new ArrayList<>();
        List<String> goodStudents = new ArrayList<>();
        List<String> satisfactoryStudents = new ArrayList<>();
        List<String> failStudents = new ArrayList<>();

        // Анализ данных
        analyzeData(sheet, excellentStudents, goodStudents, satisfactoryStudents, failStudents);

        // Создание новой книги для записи результатов
        Workbook resultWorkbook = new Workbook();
        Worksheet resultSheet = resultWorkbook.getWorksheets().get(0);

        // Запись результатов
        writeSummaryResults(resultSheet, excellentStudents, goodStudents, satisfactoryStudents, failStudents);

        logger.info("Записана информация о студентах и оценках");

        // Создаем столбчатую диаграмму
        createChart(resultSheet);

        // Сохранение результирующей книги
        resultWorkbook.save(outputFilePath, SaveFormat.XLSX);
    }

    /**
     * Анализ данных о студентах и оценках
     *
     * @param sheet                Лист с данными
     * @param excellentStudents    Список отличников
     * @param goodStudents         Список хорошистов
     * @param satisfactoryStudents Список троечников
     * @param failStudents         Список не допущенных к экзамену
     */
    private void analyzeData(Worksheet sheet, List<String> excellentStudents, List<String> goodStudents,
                             List<String> satisfactoryStudents, List<String> failStudents) {
        Cells cells = sheet.getCells();
        for (int i = 1; i < cells.getMaxDataRow() + 1; i++) {
            String studentName = cells.get(i, 0).getStringValue();
            int grade = cells.get(i, 1).getIntValue();

            if (grade == 5) {
                excellentStudents.add(studentName);
            } else if (grade == 4) {
                goodStudents.add(studentName);
            } else if (grade == 3) {
                satisfactoryStudents.add(studentName);
            } else {
                failStudents.add(studentName);
            }
        }
    }

    /**
     * Запись сводных результатов в итоговый лист
     *
     * @param resultSheet          Лист для записи результатов
     * @param excellentStudents    Список отличников
     * @param goodStudents         Список хорошистов
     * @param satisfactoryStudents Список троечников
     * @param failStudents         Список не допущенных к экзамену
     */
    private void writeSummaryResults(Worksheet resultSheet, List<String> excellentStudents, List<String> goodStudents,
                                     List<String> satisfactoryStudents, List<String> failStudents) {
        resultSheet.getCells().get("A1").setValue("Отличники");
        resultSheet.getCells().get("B1").setValue("Хорошисты");
        resultSheet.getCells().get("C1").setValue("Троешники");
        resultSheet.getCells().get("D1").setValue("Не допущены");
        resultSheet.getCells().get("E1").setValue("Средний балл");

        resultSheet.getCells().get("A2").setValue(excellentStudents.size());
        resultSheet.getCells().get("B2").setValue(goodStudents.size());
        resultSheet.getCells().get("C2").setValue(satisfactoryStudents.size());
        resultSheet.getCells().get("D2").setValue(failStudents.size());
        resultSheet.getCells().get("E2").setValue(calculateAverageGrade(excellentStudents.size(), goodStudents.size(), satisfactoryStudents.size(), failStudents.size()));

        // Запись данных о студентах в разных столбцах
        writeStudentsData(resultSheet, 6, excellentStudents, 1, "Отличники");
        writeStudentsData(resultSheet, 7, goodStudents, 1, "Хорошисты");
        writeStudentsData(resultSheet, 8, satisfactoryStudents, 1, "Троешники");
        writeStudentsData(resultSheet, 9, failStudents, 1, "Не допущены");
    }

    /**
     * Вычисление среднего балла группы
     *
     * @param excellentCount    Количество отличников
     * @param goodCount         Количество хорошистов
     * @param satisfactoryCount Количество троечников
     * @param failCount         Количество не допущенных к экзамену
     * @return Средний балл
     */
    private double calculateAverageGrade(int excellentCount, int goodCount, int satisfactoryCount, int failCount) {
        int totalStudents = excellentCount + goodCount + satisfactoryCount + failCount;
        return (double) (5 * excellentCount + 4 * goodCount + 3 * satisfactoryCount + failCount) / totalStudents;
    }

    /**
     * Запись данных о студентах в указанный столбец итогового листа
     *
     * @param sheet    Лист для записи данных
     * @param column   Номер столбца, в который будут записаны данные
     * @param students Список студентов
     * @param startRow Начальная строка для записи данных
     * @param header   Заголовок столбца
     */
    private void writeStudentsData(Worksheet sheet, int column, List<String> students, int startRow, String header) {
        sheet.getCells().get(startRow - 1, column).setValue(header);
        for (int i = 0; i < students.size(); i++) {
            sheet.getCells().get(startRow + i, column).setValue(students.get(i));
        }
    }

    /**
     * Создание диаграммы и добавление на указанный лист
     *
     * @param sheet Лист для создания диаграммы
     */
    private void createChart(Worksheet sheet) {
        int chartIndex = sheet.getCharts().add(ChartType.COLUMN, 5, 0, 15, 5);
        Chart chart = sheet.getCharts().get(chartIndex);

        chart.setType(ChartType.BAR);
        chart.setChartDataRange("A1:E2", true);

        logger.info("Создан график");
    }
}
