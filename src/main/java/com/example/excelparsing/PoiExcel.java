package com.example.excelparsing;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class PoiExcel {

    public static final String excelFilePath = "/home/ktj/바탕화면";
    public static final String excelFileName = "test_busan.xlsx";
    public static final String[ ] intents = {"구급", "구조", "화재", "기타", "추가문의", "콜센터"};
    public static final String txtFilePath = "/home/ktj/바탕화면/의도/";

    //의도 관리 List
    public static List<String> firstAidList = new ArrayList<>(); // 구급
    public static List<String> rescrueList = new ArrayList<>(); // 구조
    public static List<String> fireList = new ArrayList<>(); // 화재
    public static List<String> etcList = new ArrayList<>(); // 기타
    public static List<String> additionalInquiryList = new ArrayList<>(); // 추가문의

    //콜센터 발화 List
    public static List<String> callCenterList = new ArrayList<>(); // 콜센터

    public static void main(String args[ ]) {

        System.out.println("------parsing program running ---------");
        readIntentFiles();
        //CreateExcel();
        ReadExcel();
        System.out.println("------parsing program end ---------");
    }

    public static void readIntentFiles( )  {
        System.out.println("------readIntentFIles---------");
        try{
            for(int i=0; i<intents.length; i++) {
                //파일 객체 생성

                String FileName = txtFilePath + intents[i] + ".txt";
                File file = new File(FileName);
                //입력 스트림 생성
                FileReader filereader = new FileReader(file);
                //입력 버퍼 생성
                BufferedReader bufReader = new BufferedReader(filereader);
                String line = "";
                while ((line = bufReader.readLine()) != null) {
                    if(line == null || line.length() == 0) continue;
                    if(i==0) firstAidList.add(line);
                    else if(i==1) rescrueList.add(line);
                    else if(i==2) fireList.add(line);
                    else if(i==3) etcList.add(line);
                    else if(i==4) additionalInquiryList.add(line);
                    else if(i==5) callCenterList.add(line);
                }
                //.readLine()은 끝에 개행문자를 읽지 않는다.
                bufReader.close();
            }
        }catch (FileNotFoundException e) {
            // TODO: handle exception
        }catch(IOException e){
            System.out.println(e);
        }
    }

    public static void CreateExcel( ) {
        // 빈 Workbook 생성
        XSSFWorkbook workbook = new XSSFWorkbook();

        // 빈 Sheet를 생성
        XSSFSheet sheet = workbook.createSheet("employee data");

        // Sheet를 채우기 위한 데이터들을 Map에 저장
        Map<Integer, Object[]> data = new TreeMap<>();
        data.put(1, new Object[]{"파일명", "인식결과", "실제발화", "한글화", "출동상황", "출동상황(수정)", "비고"});
        data.put(2, new Object[]{"/20201115_000456_1017-065_split_000.wav", "(119//일일구)입니다", "(119//일일구)입니다", "(119//일일구)입니다", "해당없음", "해당없음", "신고자"});
        data.put(3, new Object[]{"/20201115_000456_1017-064_split_000.wav", "나무그늘 들어오는 거 같거든예", "나무그늘 들어오는 거 같거든예", "나무그늘 들어오는 거 같거든예", "화재", "화재", "신고자"});
        data.put(4, new Object[]{"/20201115_000456_1017-061_split_000.wav", "사람이 건물사이에 끼었어요", "사람이 건물사이에 끼었어요", "사람이 건물사이에 끼었어요", "해당없음", "해당없음", "신고자"});
        data.put(5, new Object[]{"/20201115_000456_1017-031_split_000.wav", "(119//일일구)입니다", "(119//일일구)입니다", "(119//일일구)입니다", "해당없음", "해당없음", "신고자"});

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
            FileOutputStream out = new FileOutputStream(new File(excelFilePath, excelFileName));
            workbook.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void ReadExcel() {
        System.out.println("------start ReadExcel---------");
        try {
            FileInputStream file = new FileInputStream(new File(excelFilePath, excelFileName));

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

                int idx = -1; // 출동상황중 해당없음 카테고리만 확인하면됨
                boolean mustChange=false;
                String ttsSentense = "";
                String tobeIntent = "";

                while (cellIterator.hasNext()) {
                    idx++;
                    Cell cell = cellIterator.next();
                    // cell의 타입을 하고, 값을 가져온다.
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print((int) cell.getNumericCellValue() + "\t"); //getNumericCellValue 메서드는 기본으로 double형 반환
                            break;
                        case Cell.CELL_TYPE_STRING:
                            if(idx == 3) ttsSentense = cell.getStringCellValue();
                            else if(idx == 4) {
                                String intent = cell.getStringCellValue();
                                if(intent.equals("해당없음")) {
                                    tobeIntent = matchSencenToIntent(ttsSentense);
                                    if(tobeIntent != null) {
                                        mustChange = true;
                                    }
                                }
                            } else if(idx == 5 && mustChange) {
                                cell.setCellValue(tobeIntent);
                            } else if(idx == 6) {
                                int speakerType = getSpeakerType(ttsSentense);
                                if(speakerType == 1) { // 콜센터
                                    cell.setCellValue("콜센터");
                                }
                            }
                            //System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
            }

            //수정된 열을 다시 써준다.
            FileOutputStream out = new FileOutputStream(new File(excelFilePath, excelFileName));
            workbook.write(out);
            out.close();
            file.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 추후 의도별로 쓰레드 생성해서 병렬처리 고려
    public static String matchSencenToIntent(String ttsSentense) {
        List<String> intentList = null;

        for(int i=0; i<intents.length; i++) {
            if(i==0) intentList = firstAidList;
            else if(i==1) intentList = rescrueList;
            else if(i==2) intentList = fireList;
            else if(i==3) intentList = etcList;
            else if(i==4) intentList = additionalInquiryList;

            for(int j=0; j<intentList.size(); j++) {
                String learningSentense = intentList.get(j);

                //ttsSentense가 학습데이터를 포함하면
                boolean isContained = ttsSentense.contains(learningSentense);
                if(isContained) return intents[i];
            }
        }
        return null;
    }

    // 콜센터 or 신고자
    // 콜센터 : 1
    // 신고자 : 2
    public static int getSpeakerType(String ttsSentense) {

        for(int i=0; i<callCenterList.size(); i++) {
            //ttsSentense가 학습데이터를 포함하면
            String learningSentense = callCenterList.get(i);
            boolean isContained = ttsSentense.contains(learningSentense);
            if(isContained) return 1;
        }

        return 2;
    }
}
