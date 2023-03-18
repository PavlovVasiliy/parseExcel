package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ParseBP {
    static int startPos = 2; // Начало данных с 3 строки в файле
    static Workbook resultFile = Main.resultFile;

    static Sheet sheetModel = resultFile.getSheet("БП");

    private static Map<Integer, List<String>> defineLimitsMap() {

        Map<Integer, List<String>> map = new HashMap<>();
        Integer pos = 0;

        // Берем лист с бизнес процессами и выбираем заголовки
        Row interfaceHeaders = sheetModel.getRow(1);

        // Заполняем карту на кажду позицию шаблона
        for (int i = 0; i < interfaceHeaders.getLastCellNum() - 2; i++) {
            List<String> fields = new ArrayList<>();
            fields.add(interfaceHeaders.getCell(i).getStringCellValue().trim().toLowerCase());

            // 0 - варианты столбца А
            if (i == 0) {
                fields.add(null);
            }

            // 1 - варианты столбца B
            if (i == 1) {
                fields.add("этап id".toLowerCase());
                fields.add("ID".toLowerCase());
                fields.add("Шаг".toLowerCase());
                fields.add("Номер этапа".toLowerCase());
                fields.add("IDшага".toLowerCase());
            }

            // 2 - варианты столбца C
            if (i == 2) {
                fields.add("Имя".toLowerCase());
                fields.add("Название".toLowerCase());
                fields.add("Наименование.1".toLowerCase());
            }

            // 3 - варианты столбца D
            if (i == 3) {

            }

            // 4 - варианты столбца E
            if (i == 4) {
                fields.add("вход".toLowerCase());
                fields.add("Входные даные".toLowerCase());
            }

            // 5 - варианты столбца F
            if (i == 5) {
                fields.add("выход".toLowerCase());
            }

            // 6 - варианты столбца G
            if (i == 6) {
                fields.add("Роль".toLowerCase());
                fields.add("IDроли".toLowerCase());
            }

            // 7 - варианты столбца H
            if (i == 7) {
                fields.add("документ".toLowerCase());
            }

            // 8 - варианты столбца I
            if (i == 8) {
                fields.add("система-обработчик".toLowerCase());
                fields.add("Соответствие Системы".toLowerCase());
                fields.add("Мэппинг с системой".toLowerCase());
            }

            // 9 - варианты столбца J
            if (i == 9) {
                fields.add("поддержка системы".toLowerCase());
                fields.add("Поддержкасистемы".toLowerCase());
            }

            map.put(pos++, fields);
        }

        return map;
    }

    public static void parseBP(Workbook curWorkbook, String fileName) throws IOException {

        // Лист Sheet'ов которые удоволетворяют условию
        List<String> curListNameSheet = new ArrayList<>();

        // Мапа для определения листа с бизнес процессами
        Map<Integer, List<String>> limitsMap = defineLimitsMap();

        // Проход по листам
        for (int i = 0; i < curWorkbook.getNumberOfSheets(); i++) {
            Sheet curSheet = curWorkbook.getSheetAt(i);
            Row curHeaders = curSheet.getRow(0);

            // Собираем лист текущих заголовков столбцов
            List<String> listCurFields = getHeadersList(curHeaders);

            // Определяем является ли данный лист листом с бизнес процессами
            boolean isConditionOk = isConditionSheetOk(listCurFields, limitsMap);

            // Если лист с бизнес процессами
            if (isConditionOk) {
                curListNameSheet.add(curSheet.getSheetName());

                // Копирование строк
                copyRows(curSheet, fileName, limitsMap, listCurFields);
            }
        }

        if (curListNameSheet.size() == 0)
            System.out.println("В файле " + fileName + " бизнес процессов не найдено");
        else
            System.out.println("В файле " + fileName + " бизнес процессам соответствуют листы: " + curListNameSheet);
    }


    private static void copyRows(Sheet curSheet, String fileName, Map<Integer, List<String>> map, List<String> listCurFields) {
        for (int i = 1; i <= curSheet.getLastRowNum(); i++) {
            Row row = curSheet.getRow(i);
            if (row == null) continue;

            while (row.getLastCellNum() < map.size())
                row.createCell(row.getLastCellNum());

            //Выравнивание строки начиная с начала
            int indexCurSheetHeaderList = 0;
            for (Integer j = 0; j < map.size(); j++) {
                if (indexCurSheetHeaderList >= listCurFields.size() || !map.get(j).contains(listCurFields.get(indexCurSheetHeaderList++))) {
                    indexCurSheetHeaderList--;
                }
            }

            // Проверка, если слишком длинный
            while (row.getLastCellNum() > map.size()) {
                row.removeCell(row.getCell(row.getLastCellNum() - 1));
            }

            // Добавить в ячейки имя листа и имя файла
            addSheetNameAndFileNameToEndOfRow(row, curSheet.getSheetName(), fileName);

            // Добавление новой строки к текущему рабочему листу
            copyRowToCur(row);

        }
    }

    private static boolean isConditionSheetOk(List<String> listCurFields, Map<Integer, List<String>> map) {

        boolean isConditionOk = true;

        int count = 0;
        // Проверка по каждой корзине
        for (Integer i = 0; i < listCurFields.size() && i < map.size(); i++) {
            count++;
            if (!map.get(i).contains(listCurFields.get(i))) {
                if (count < 5) {
                    isConditionOk = false;
                    break;
                }
            }

        }

        return isConditionOk;
    }

    private static List<String> getHeadersList(Row headers) {
        List<String> listCurFields = new ArrayList<>();
        for (int i = 0; i < headers.getLastCellNum(); i++) {
            listCurFields.add((headers.getCell(i) == null) ?
                    null : headers.getCell(i).getStringCellValue().trim().toLowerCase());
        }
        return listCurFields;
    }

    private static void addSheetNameAndFileNameToEndOfRow(Row row, String sheetName, String fileName) {
        row.createCell(row.getLastCellNum()).setCellValue(sheetName);
        row.createCell(row.getLastCellNum()).setCellValue(fileName);
    }

    private static void copyRowToCur(Row row) {
        Row rowNew = sheetModel.createRow(startPos++);

        for (int j = 0; j <= row.getLastCellNum(); j++) {
            if (row.getCell(j) == null) continue;
            switch (row.getCell(j).getCellType()) {
                case _NONE:
                    break;
                case NUMERIC:
                    rowNew.createCell(j).setCellValue(row.getCell(j).getNumericCellValue());
                    break;
                case STRING:
                    rowNew.createCell(j).setCellValue(row.getCell(j).getStringCellValue());
                    break;
                default:
                    rowNew.createCell(j).setCellValue("");
            }
        }
    }
}
