package org.example;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.List;

public class Main {

    private static String dirPath = "C:\\Main\\Северсталь\\Модели\\От северстали\\QMET_USERS_BP_cur\\";
    private static String resultFileName = "result\\Результат_в3.xlsx";

    public static List<String> excelList = Excel.getListExcelFilesInDir(dirPath);

    public static Workbook resultFile = Excel.openWorkBook(dirPath, resultFileName);

    static List<String> listExcelFilesInDir = Excel.getListExcelFilesInDir(dirPath);

    public static void main(String[] args) throws IOException {


        // Проход по всем файлам
        for (String fileName : listExcelFilesInDir) {
//            ParseRoles.parseRoles(Excel.openWorkBook(dirPath, fileName), fileName);
//            ParseInterfaces.parseInterfacies(Excel.openWorkBook(dirPath, fileName), fileName);
//            ParseObjects.parseObjects(Excel.openWorkBook(dirPath, fileName), fileName);


        }

        // Запись результата
        Excel.saveWorkBook(resultFile, dirPath, resultFileName);


    }
}