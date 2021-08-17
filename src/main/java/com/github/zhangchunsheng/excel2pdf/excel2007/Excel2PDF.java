package com.github.zhangchunsheng.excel2pdf.excel2007;

import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Text;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @date 2021/8/17
 */
public class Excel2PDF {

    private XSSFSheet sheet;

    private Document document;

    private float rate;

    private float excelWidth;

    private int lastCellNum;

    private String fontPath;

    public Excel2PDF(InputStream inputStream) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        this.sheet = workbook.getSheetAt(0);
    }

    public Excel2PDF(InputStream inputStream, OutputStream outputStream) throws IOException {
        this(inputStream);
        PdfDocument pdfDocument = new PdfDocument(new PdfWriter(outputStream));
        this.document = new Document(pdfDocument, PageSize.A4.rotate());
        this.rate = getRate();
        this.lastCellNum = this.sheet.getRow(0).getLastCellNum();
    }

    public Excel2PDF(InputStream inputStream, OutputStream outputStream, String fontPath) throws IOException {
        this(inputStream);
        PdfDocument pdfDocument = new PdfDocument(new PdfWriter(outputStream));
        this.document = new Document(pdfDocument, PageSize.A4.rotate());
        this.rate = getRate();
        this.lastCellNum = this.sheet.getRow(0).getLastCellNum();
        this.fontPath = fontPath;
    }

    /**
     * 转换
     *
     * @throws IOException
     */
    public void convert() throws IOException {
        Table table = new Table(getColumnWidths());
        doRowAndCell(table);
        doPicture(table);
        document.add(table);
        document.close();
    }

    /**
     * 处理图片
     *
     * @param table
     */
    private void doPicture(Table table) {
        XSSFDrawing dp = (XSSFDrawing) sheet.createDrawingPatriarch();
        if(dp != null){
            List<XSSFShape> children = dp.getShapes();
            for (XSSFShape shape : children){
                XSSFPicture xssfPicture = (XSSFPicture)shape;
                table.setNextRenderer(new OverlappingImageTableRenderer(table, xssfPicture, sheet));
            }
        }
    }

    /**
     * 处理行列
     *
     * @param table
     * @throws IOException
     */
    private void doRowAndCell(Table table) throws IOException {
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i < lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int j = 0; j < lastCellNum; j++) {
                XSSFCell cell = row.getCell(j);
                if (cell != null) {
                    if (!isUse(cell)) {
                        CellRangeAddress cellRangeAddress = getCellRangeAddress(cell);
                        int rowspan = 1;
                        int colspan = 1;
                        if (cellRangeAddress != null) {
                            colspan = cellRangeAddress.getLastColumn() - cellRangeAddress.getFirstColumn() + 1;
                            rowspan = cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow() + 1;
                            j = cellRangeAddress.getLastColumn();
                        }
                        Cell pdfCell = transformCommon(cell, rowspan, colspan);
                        table.addCell(pdfCell);
                    }
                } else {
                    // 补偿
                    table.addCell("");
                }
            }
        }
    }

    /**
     * 处理每一个单元格
     *
     * @param cell
     * @param rowspan
     * @param colspan
     * @return
     * @throws IOException
     */
    private Cell transformCommon(XSSFCell cell, int rowspan, int colspan) throws IOException {
        String value = Excel2PdfHelper.getValue(cell);

        Cell pdfCell = new Cell(rowspan, colspan)
                .setHeight(cell.getRow().getHeight() * this.rate * 1.2f)
                .setPadding(1);
        Text text = new Text(value);
        setPdfCellFont(cell, text);
        pdfCell.add(new Paragraph(text).setPadding(0f).setMargin(0f));
        XSSFCellStyle cellStyle = cell.getCellStyle();
        // 布局
        VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();
        pdfCell.setVerticalAlignment(Excel2PdfHelper.getVerticalAlignment(verticalAlignment));
        HorizontalAlignment alignment = cellStyle.getAlignment();
        pdfCell.setTextAlignment(Excel2PdfHelper.getTextAlignment(alignment, cell.getCellType()));

        //边框
        Excel2PdfHelper.transformBorder(cell, pdfCell);

        //背景色
        XSSFColor xSSFColor = cellStyle.getFillBackgroundColorColor();

        if (xSSFColor != null) {
            byte[] rgb = xSSFColor.getRGB();
            if(rgb != null) {
                pdfCell.setBackgroundColor(new DeviceRgb(Byte.toUnsignedInt(rgb[0]), Byte.toUnsignedInt(rgb[1]), Byte.toUnsignedInt(rgb[2])));
            }
        }
        return pdfCell;
    }

    /**
     * 处理单元格字体样式
     *
     * @param cell
     * @param text
     * @throws IOException
     */
    private void setPdfCellFont(XSSFCell cell, Text text) throws IOException {
        XSSFCellStyle cellStyle = cell.getCellStyle();
        //字体大小
        XSSFFont font = cellStyle.getFont();
        short fontHeight = font.getFontHeight();
        if(this.fontPath != null && !this.fontPath.equals("")) {
            text.setFont(PdfFontFactory.createFont(this.fontPath, PdfEncodings.IDENTITY_H));
        } else {
            text.setFont(PdfFontFactory.createFont(System.getProperty("user.dir") + "/doc/font/SimHei.TTF", PdfEncodings.IDENTITY_H));
        }

        text.setFontSize(fontHeight * rate * 1.05f);

        //字体颜色
        XSSFColor xssfColor = font.getXSSFColor();
        if (xssfColor != null && xssfColor.getIndex() != 64) {
            byte[] rgb = xssfColor.getRGB();
            if(rgb != null) {
                text.setFontColor(new DeviceRgb(Byte.toUnsignedInt(rgb[0]), Byte.toUnsignedInt(rgb[1]), Byte.toUnsignedInt(rgb[2])));
            }
        }

        //加粗
        if (font.getBold()) {
            text.setBold();
        }

        // 斜体
        if (font.getItalic()) {
            text.setItalic();
        }

        // 下划线
        if (font.getUnderline() == 1) {
            text.setUnderline(0.5f, -1f);
        }
    }

    private XSSFPicture getXSSFPicture(HSSFCell cell) {
        XSSFDrawing dp = sheet.getDrawingPatriarch();
        if (dp != null) {
            List<XSSFShape> children = dp.getShapes();
            for (XSSFShape shape : children) {
                XSSFPicture xssfPicture = (XSSFPicture) shape;
                XSSFClientAnchor clientAnchor = xssfPicture.getClientAnchor();
                if (cell.getRowIndex() == clientAnchor.getRow1() && cell.getColumnIndex() == clientAnchor.getCol1()) {
                    return xssfPicture;
                }
            }
        }
        return null;
    }

    /**
     * 获取PDF纸张和Excel总宽度的比值
     * 用这个值对Excel的大小缩放成pdf需要的大小
     *
     * @return
     */
    private float getRate() {
        float all = 0;
        short lastCellNum = this.sheet.getRow(0).getLastCellNum();
        for (int i = 0; i < lastCellNum; i++) {
            all += this.sheet.getColumnWidth(i);
        }
        PageSize defaultPageSize = null;
        if (document == null) {
            defaultPageSize = PageSize.Default;
        } else {
            defaultPageSize = document.getPdfDocument().getDefaultPageSize();
        }
        this.excelWidth = all;
        float width = defaultPageSize.getWidth();
        return width / all;
    }

    /**
     * 获取单元格列宽
     *
     * @return
     */
    private float[] getColumnWidths() {
        float[] widths = new float[lastCellNum];
        for (int i = 0; i < lastCellNum; i++) {
            int columnWidth = this.sheet.getColumnWidth(i);
            float realWidth = columnWidth * rate;
            widths[i] = realWidth;
        }
        return widths;
    }

    /**
     * 判断单元格是否处理过
     * 合并的单元格只处理第一个
     *
     * @param cell
     * @return
     */
    private boolean isUse(XSSFCell cell) {
        List<CellRangeAddress> mergedRegions = cell.getSheet().getMergedRegions();
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        for (CellRangeAddress cellAddresses : mergedRegions) {
            if (cellAddresses.getFirstColumn() <= columnIndex && cellAddresses.getLastColumn() >= columnIndex
                    && cellAddresses.getFirstRow() <= rowIndex && cellAddresses.getLastRow() >= rowIndex
                    && !(cellAddresses.getFirstRow() == rowIndex && cellAddresses.getFirstColumn() == columnIndex)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 获取合并单元格，只处理第一个
     *
     * @param cell
     * @return
     */
    private CellRangeAddress getCellRangeAddress(XSSFCell cell) {
        List<CellRangeAddress> mergedRegions = cell.getSheet().getMergedRegions();
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        for (CellRangeAddress cellAddresses : mergedRegions) {
            if (cellAddresses.getFirstRow() == rowIndex && cellAddresses.getFirstColumn() == columnIndex) {
                return cellAddresses;
            }
        }
        return null;
    }
}
