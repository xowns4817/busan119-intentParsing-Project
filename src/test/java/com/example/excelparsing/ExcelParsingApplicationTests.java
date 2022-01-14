
package com.example.excelparsing;

;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

public class ExcelParsingApplicationTests {

    //의도목록 메모리 로드
    @BeforeAll
    public static void intentSetting( ) {
        System.out.println("ok!");
        PoiExcel.printInitLog();
        PoiExcel.readIntentFiles();
    }

    // 의도 Test
    @Test
    public void intentTest( ) {
        String tts = "선생님 안녕하세요. 불난곳이 거기가 어디 입니까?";
        int intentType = PoiExcel.matchSencenToIntent(tts);
        assertThat(intentType).isEqualTo(-1);
    }

    // 콜센터/신고자 Test
    @Test
    public void etcTest( ) {

        String tts = "선생님 안녕하세요. 거기가 어디 입니까?";
        int speakerType = PoiExcel.getSpeakerType(tts);

        assertThat(speakerType).isEqualTo(0);
    }

    @AfterAll
    public static void releaseResource( ) {
        System.out.println("ok!");
    }
}