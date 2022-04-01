package dto;

import annotation.AllCell;
import annotation.Cell;

import java.time.LocalDateTime;

@AllCell
public class Dummy {
    @Cell(headerName = "제목")
    private String title;

    @Cell(headerName = "내용")
    private String contents;

    @Cell(headerName = "날짜/시간")
    private LocalDateTime dateTime;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getContents() {
        return contents;
    }

    public void setContents(String contents) {
        this.contents = contents;
    }

    public LocalDateTime getDateTime() {
        return dateTime;
    }

    public void setDateTime(LocalDateTime dateTime) {
        this.dateTime = dateTime;
    }

    public Dummy(String title, String contents, LocalDateTime dateTime) {
        this.title = title;
        this.contents = contents;
        this.dateTime = dateTime;
    }
}
