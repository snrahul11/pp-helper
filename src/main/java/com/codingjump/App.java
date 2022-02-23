package com.codingjump;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class App {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        ppt.createSlide();
        FileOutputStream out = new FileOutputStream("./powerpoint.pptx");
        ppt.write(out);
        out.close();
    }
}
