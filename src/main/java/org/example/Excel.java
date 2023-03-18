package org.example;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class Excel {
    public static void main(String[] args) throws IOException {

        String dirPath = "C:\\Main\\Северсталь\\Модели\\От северстали\\QMET_USERS_BP_cur\\";
//        String filePath = "CherMK_BPD_Severstal_CS_v1.0_RUS.docx.xlsx";


        //define excel list
        List<String> excelList = getListExcelFilesInDir(dirPath);

        // AutoFit All files
        for (String fileName : excelList) {
            //read
            Workbook workbook = openWorkBook(dirPath, fileName);


            //autoFit
            autoFitColumns(workbook);

            //write
            saveWorkBook(workbook, dirPath, fileName);
        }
    }

    protected static Workbook openWorkBook(String dir, String file) {
        try (FileInputStream fileInputStream = new FileInputStream(dir + file)) {
            return new XSSFWorkbook(new FileInputStream(dir + file));
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage() + "\nфайл " + file + " не найден!");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return null;
    }

    protected static void saveWorkBook(Workbook workbook, String dir, String file) {
        try (FileOutputStream fileOutputStream = new FileOutputStream(dir + file)) {
            workbook.write(fileOutputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage() + "\nфайл " + file + " не найден!");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static void autoFitColumns(Workbook workbook) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            for (int j = 0; j < sheet.getRow(0).getLastCellNum(); j++) {
                sheet.autoSizeColumn(j);
            }
        }
    }

    protected static List<String> getListExcelFilesInDir(final String pathFolder) {
        File dir = new File(pathFolder);
        List<String> list = new ArrayList<>();

        if (dir.isDirectory()) {
            for (final File fileEntry : dir.listFiles()) {
                String fileName = fileEntry.getName();
                if (fileEntry.isFile() && fileName.endsWith(".xlsx"))
                    list.add(fileName);
            }
        }
        Collections.sort(list);
        return list;
    }
}
