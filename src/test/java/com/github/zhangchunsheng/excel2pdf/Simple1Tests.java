package com.github.zhangchunsheng.excel2pdf;

import com.github.zhangchunsheng.excel2pdf.excel2003.Excel2PDF;
import org.junit.Test;

import java.io.*;
import java.net.URL;
import java.util.UUID;

public class Simple1Tests {
    @Test
    public void testCase1OfSingle() throws IOException {
        URL url = this.getClass().getResource("sample1/case1.xls");
        url = this.getClass().getResource("sample1/case5.xlsx");
        url = this.getClass().getResource("sample1/case1.xlsx");

        System.out.println(url.getPath());

        Excel2PdfTool excel2PdfTool = new Excel2PdfTool();
        excel2PdfTool.convertToPdf(url.getPath(), "output1.pdf");
    }

    @Test
    public void test() throws IOException {
        InputStream inputStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("sample1/case1.xls");
        inputStream = this.getClass().getResourceAsStream("sample1/case1.xlsx");
        String name = UUID.randomUUID().toString().substring(0, 10);
        name = "output1";
        FileOutputStream outputStream = new FileOutputStream(name + ".pdf");
        Excel2PDF excel2PDF = new Excel2PDF(inputStream, outputStream);
        excel2PDF.convert();
    }

    private File fileOut(String fileIn) {
        String uri = this.getClass().getResource(fileIn).getPath();
        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
        File file = new File(fileOut);
        return file;
    }
}