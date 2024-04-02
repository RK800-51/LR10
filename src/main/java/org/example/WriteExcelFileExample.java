package org.example;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class WriteExcelFileExample {
    public static void main(String[] args) throws  IOException{
        // Создаем новый файл excel
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Создаем новый лист в файле
        XSSFSheet sheet = workbook.createSheet("Товары");

        // Записываем данные в ячейки
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Товар");
        headerRow.createCell(1).setCellValue("Характеристики");
        headerRow.createCell(2).setCellValue("Стоимость");

        Row dataRow1 = sheet.createRow(1);
        dataRow1.createCell(0).setCellValue("Книга");
        dataRow1.createCell(1).setCellValue("Жанр: Фантастика, Автор: Иванов И.И.");
        dataRow1.createCell(2).setCellValue((500.0));

        Row dataRow2 = sheet.createRow(2);
        dataRow2.createCell(0).setCellValue("Процессор");
        dataRow2.createCell(1).setCellValue("Intel Core i5");
        dataRow2.createCell(2).setCellValue((25000.0));

        // Записываем объект Excel в файл
        String filePath = "src/main/java/org/example/example-excel.xlsx";
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();

        System.out.println("Excel file created: " + filePath);

    }
}
