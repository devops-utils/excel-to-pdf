package com.github.zhangchunsheng.excel2pdf;

import com.github.zhangchunsheng.excel2pdf.excel2003.Excel2PDF;

import java.io.*;

public class EPFactory {
    private final static String EXCEL2003 = "xls";
    private final static String EXCEL2007 = "xlsx";

    public static IExcel2PDF getEP(String xlsPath, String pdfPath) throws IOException {
        if (xlsPath.endsWith(EXCEL2007)) {
            InputStream inputStream = new FileInputStream(xlsPath);
            FileOutputStream outputStream = new FileOutputStream(pdfPath);
            return new com.github.zhangchunsheng.excel2pdf.excel2007.Excel2PDF(inputStream, outputStream);
        } else if (xlsPath.endsWith(EXCEL2003)) {
            InputStream inputStream = new FileInputStream(xlsPath);
            FileOutputStream outputStream = new FileOutputStream(pdfPath);
            return new Excel2PDF(inputStream, outputStream);
        }
        return null;
    }

    public static IExcel2PDF getEP(String xlsPath, String pdfPath, String fontPath) throws IOException {
        if (xlsPath.endsWith(EXCEL2007)) {
            InputStream inputStream = new FileInputStream(xlsPath);
            FileOutputStream outputStream = new FileOutputStream(pdfPath);
            return new com.github.zhangchunsheng.excel2pdf.excel2007.Excel2PDF(inputStream, outputStream, fontPath);
        } else if (xlsPath.endsWith(EXCEL2003)) {
            InputStream inputStream = new FileInputStream(xlsPath);
            FileOutputStream outputStream = new FileOutputStream(pdfPath);
            return new Excel2PDF(inputStream, outputStream, fontPath);
        }
        return null;
    }
}
