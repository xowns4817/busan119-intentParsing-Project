package com.example.excelparsing;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class PoiReadExcel {
    public static String filePath = "/home/ktj/바탕화면";
    public static String fileNm = "poi_making_file_test.xlsx";

    public static void main(String[] args) {

        try {
            FileInputStream file = new FileInputStream(new File(filePath, fileNm));

            // 엑셀 파일로 Workbook instance를 생성한다.
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // workbook의 첫번째 sheet를 가저온다.
            XSSFSheet sheet = workbook.getSheetAt(0);

            // 만약 특정 이름의 시트를 찾는다면 workbook.getSheet("찾는 시트의 이름");

            // 모든 행(row)들을 조회한다.
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                org.apache.poi.ss.usermodel.Row row = rowIterator.next();

                // 각각의 행에 존재하는 모든 열(cell)을 순회한다.
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    // cell의 타입을 하고, 값을 가져온다.
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print((int) cell.getNumericCellValue() + "\t"); //getNumericCellValue 메서드는 기본으로 double형 반환
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println();
            }
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
