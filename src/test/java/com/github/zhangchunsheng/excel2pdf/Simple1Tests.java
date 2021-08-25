package com.github.zhangchunsheng.excel2pdf;

import com.github.zhangchunsheng.excel2pdf.excel2003.Excel2PDF;
import org.junit.Test;

import java.io.*;
import java.net.URL;
import java.util.UUID;

/**
 * <pre>
 * Created by Chunsheng Zhang on 2020/8/17.
 * </pre>
 *
 * @author <a href="https://github.com/zhangchunsheng">Chunsheng Zhang</a>
 */
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
        inputStream = this.getClass().getResourceAsStream("sample1/case1.xls");
        String name = UUID.randomUUID().toString().substring(0, 10);
        name = "output1";
        FileOutputStream outputStream = new FileOutputStream(name + ".pdf");
        Excel2PDF excel2PDF = new Excel2PDF(inputStream, outputStream);
        excel2PDF.convert();
    }

    @Test
    public void testXlsx() throws IOException {
        InputStream inputStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("sample1/case1.xls");
        inputStream = this.getClass().getResourceAsStream("sample1/case1.xlsx");
        String name = UUID.randomUUID().toString().substring(0, 10);
        name = "output1";
        FileOutputStream outputStream = new FileOutputStream(name + ".pdf");
        com.github.zhangchunsheng.excel2pdf.excel2007.Excel2PDF excel2PDF = new com.github.zhangchunsheng.excel2pdf.excel2007.Excel2PDF(inputStream, outputStream);
        excel2PDF.convert();
    }

    @Test
    public void testEP() throws IOException {
        URL url = this.getClass().getResource("sample1/case1.xls");
        url = this.getClass().getResource("sample1/case5.xlsx");
        url = this.getClass().getResource("sample1/case6.xls");

        System.out.println(url.getPath());

        IExcel2PDF excel2PdfTool = EPFactory.getEP(url.getPath(), "output1.pdf", System.getProperty("user.dir") + "/doc/font/SimHei.TTF");
        if(excel2PdfTool != null) {
            excel2PdfTool.convert();
        }
    }

    private File fileOut(String fileIn) {
        String uri = this.getClass().getResource(fileIn).getPath();
        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
        File file = new File(fileOut);
        return file;
    }
}