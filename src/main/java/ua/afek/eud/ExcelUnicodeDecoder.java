package ua.afek.eud;

import org.apache.commons.lang.StringEscapeUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUnicodeDecoder {

    public ExcelUnicodeDecoder(File file, String sheetName) {
        if (!file.exists()) {
            System.out.println("File doesn't exist");
            return;
        }

        Workbook workbook = this.getWorkbook(file);
        if (workbook == null) {
            System.out.println("Error find workbook");
            return;
        }

        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            System.out.println("Error find sheet");
            return;
        }

        int count = 0;
        for (Row row : sheet)
            for (Cell cell : row)
                if (cell != null && cell.getCellType() == CellType.STRING && !cell.getStringCellValue().isEmpty()) {
                    cell.setCellValue(StringEscapeUtils.unescapeJava(cell.getStringCellValue()));
                    count++;
                }


        try {
            FileOutputStream fileOut = new FileOutputStream(file);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (IOException e) {
            System.out.println("Error save data: " + e.getMessage());
            return;
        }

        System.out.println("String decode count: " + count);
    }

    public static void main(String[] args) {
        if (args.length != 2) {
            System.out.println("Error args length");
            return;
        }

        File file = new File(args[0]);
        new ExcelUnicodeDecoder(file, args[1]);
    }

    private Workbook getWorkbook(File file) {
        try {
            String excelFilePath = file.getPath();
            FileInputStream inputStream = new FileInputStream(file);
            Workbook workbook;

            if (excelFilePath.endsWith("xlsx")) workbook = new XSSFWorkbook(inputStream);
            else if (excelFilePath.endsWith("xls"))
                workbook = new HSSFWorkbook(inputStream);
            else
                throw new IllegalArgumentException("The specified file is not Excel file");

            return workbook;

        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }
}
