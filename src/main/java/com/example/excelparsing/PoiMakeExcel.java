package com.example.excelparsing;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class PoiMakeExcel {
    public static String filePath = "/home/ktj/바탕화면";
    public static String fileNm = "poi_making_file_test.xlsx";

    public static void main(String[] args) {

        // 빈 Workbook 생성
        XSSFWorkbook workbook = new XSSFWorkbook();

        // 빈 Sheet를 생성
        XSSFSheet sheet = workbook.createSheet("employee data");

        // Sheet를 채우기 위한 데이터들을 Map에 저장
        Map<Integer, Object[]> data = new TreeMap<>();
        data.put(1, new Object[]{"파일명", "인식결과", "실제발화", "한글화", "출동상황", "출동상황(수정)"});
        data.put(2, new Object[]{"/20201115_000456_1017-065_split_000.wav", "(119//일일구)입니다", "(119//일일구)입니다", "(119//일일구)입니다", "해당없음", "해당없음"});
        data.put(3, new Object[]{"/20201115_000456_1017-064_split_000.wav", "나무그늘 들어오는 거 같거든예", "나무그늘 들어오는 거 같거든예", "나무그늘 들어오는 거 같거든예", "화재", "화재"});

        // data에서 keySet를 가져온다. 이 Set 값들을 조회하면서 데이터들을 sheet에 입력한다.
        Set<Integer> keyset = data.keySet();
        int rownum = 0;

        // 알아야할 점, TreeMap을 통해 생성된 keySet는 for를 조회시, 키값이 오름차순으로 조회된다.
        for (Integer key : keyset) {
            XSSFRow row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String) {
                    cell.setCellValue((String)obj);
                } else if (obj instanceof Integer) {
                    cell.setCellValue((Integer)obj);
                }
            }
        }

        try {
            FileOutputStream out = new FileOutputStream(new File(filePath, fileNm));
            workbook.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
