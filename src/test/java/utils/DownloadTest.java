package utils;

import dto.Dummy;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

public class DownloadTest {

    private final ExcelDownloader excelDownloader = new ExcelDownloader();
    private final List<Dummy> list = new ArrayList<>();

    @BeforeEach
    public void setUp() {
        list.add(new Dummy("Title_1", "Contents_1", LocalDateTime.now().minusDays(5L)));
        list.add(new Dummy("Title_2", "Contents_2", LocalDateTime.now().minusDays(4L)));
        list.add(new Dummy("Title_3", "Contents_3", LocalDateTime.now().minusDays(3L)));
        list.add(new Dummy("Title_4", "Contents_4", LocalDateTime.now().minusDays(2L)));
        list.add(new Dummy("Title_5", "Contents_5", LocalDateTime.now().minusDays(1L)));
        list.add(new Dummy("Title_6", "Contents_6", LocalDateTime.now()));
    }

    @Test
    public void downloadTest() {
        Assertions.assertDoesNotThrow(() -> {
            File file = new File("/Users/trenbe/Documents/JExcelDownloader-master/upload/dto.Dummy.xlsx");
            FileOutputStream stream = new FileOutputStream(file);
            excelDownloader.excelDownload(list, Dummy.class, stream, false);
        });
    }
}
