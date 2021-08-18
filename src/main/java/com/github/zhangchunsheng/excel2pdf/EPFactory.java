package com.github.zhangchunsheng.excel2pdf;

import com.github.zhangchunsheng.excel2pdf.excel2003.Excel2PDF;

import java.io.*;

/**
 * <pre>
 * Created by Chunsheng Zhang on 2020/8/17.
 * </pre>
 *
 * @author <a href="https://github.com/zhangchunsheng">Chunsheng Zhang</a>
 */
public class EPFactory {
    private final static String EXCEL2003 = "xls";
    private final static String EXCEL2007 = "xlsx";

    /**
     *
     * @param xlsPath xls path
     * @param pdfPath pdf path
     * @return An Object for excel to pdf
     * @throws IOException
     */
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

    /**
     *
     * @param xlsPath xls path
     * @param pdfPath pdf path
     * @param fontPath font path
     * @return An Object for excel to pdf
     * @throws IOException
     */
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
