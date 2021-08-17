package com.github.zhangchunsheng.excel2pdf;

import org.junit.Test;

import java.io.*;
import java.net.URL;

public class Simple1Tests {
    @Test
    public void testCase1OfSingle() throws IOException {
        URL url = this.getClass().getResource("sample1/case1.xls");
        url = this.getClass().getResource("sample1/case5.xlsx");
        url = this.getClass().getResource("sample1/case1.xlsx");

        System.out.println(url.getPath());

        App app = new App();
        app.convertToPdf(url.getPath(), "output1.pdf");
    }

    private File fileOut(String fileIn) {
        String uri = this.getClass().getResource(fileIn).getPath();
        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
        File file = new File(fileOut);
        return file;
    }
}