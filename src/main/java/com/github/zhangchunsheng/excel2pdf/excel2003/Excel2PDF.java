package com.github.zhangchunsheng.excel2pdf.excel2003;

import com.github.zhangchunsheng.excel2pdf.IExcel2PDF;
import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDictionary;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfString;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.annot.*;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Text;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.awt.*;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * <pre>
 * Created by Chunsheng Zhang on 2020/8/17.
 * </pre>
 *
 * @author <a href="https://github.com/zhangchunsheng">Chunsheng Zhang</a>
 */
public class Excel2PDF implements IExcel2PDF {

    private HSSFSheet sheet;

    private HSSFPalette customPalette;

    private PdfDocument pdfDocument;

    private Document document;

    private float rate;

    private float excelWidth;

    private int lastCellNum;

    private String fontPath;

    private Map<String, Cell> annotationsCellMap;

    private float[] columnWidths;

    public Excel2PDF(InputStream inputStream) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        this.sheet = workbook.getSheetAt(0);
        this.customPalette = workbook.getCustomPalette();
    }

    public Excel2PDF(InputStream inputStream, OutputStream outputStream) throws IOException {
        this(inputStream);
        PdfDocument pdfDocument = new PdfDocument(new PdfWriter(outputStream));
        this.pdfDocument = pdfDocument;
        this.annotationsCellMap = new HashMap<>();
        this.document = new Document(pdfDocument, PageSize.A4.rotate());
        this.rate = getRate();
        this.lastCellNum = this.sheet.getRow(0).getLastCellNum();
    }

    public Excel2PDF(InputStream inputStream, OutputStream outputStream, String fontPath) throws IOException {
        this(inputStream);
        PdfDocument pdfDocument = new PdfDocument(new PdfWriter(outputStream));
        this.pdfDocument = pdfDocument;
        this.annotationsCellMap = new HashMap<>();
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
        this.columnWidths = this.getColumnWidths();
        Table table = new Table(this.columnWidths);
        doRowAndCell(table);
        doPicture(table);
        doAnnotation(table);
        document.add(table);

        document.close();
    }

    private void doAnnotation(Table table) {
        table.setNextRenderer(new OverlappingAnnotationTableRenderer(table, annotationsCellMap, pdfDocument));
    }

    /**
     * 处理图片
     *
     * @param table
     */
    private void doPicture(Table table) {
        HSSFPatriarch drawingPatriarch = sheet.getDrawingPatriarch();
        if(drawingPatriarch != null) {
            List<HSSFShape> children = drawingPatriarch.getChildren();
            List<HSSFPicture> hSSFPictures = new ArrayList<>();
            HSSFPicture hssfPicture;
            for (HSSFShape shape : children){
                hssfPicture = (HSSFPicture)shape;
                hSSFPictures.add(hssfPicture);
            }
            table.setNextRenderer(new OverlappingImageTableRenderer(table, hSSFPictures, sheet));
        }

    }

    /**
     * 处理行列
     *
     * @param table
     * @throws IOException
     */
    private void doRowAndCell(Table table) throws IOException {
        int lastRowNum = sheet.getLastRowNum() + 1;
        int rowspan;
        int colspan;
        float maxWidth;
        for (int i = 0; i < lastRowNum; i++) {
            HSSFRow row = sheet.getRow(i);
            if(row == null) {
                for (int j = 0; j < lastCellNum; j++) {
                    Cell pdfCell = new Cell();
                    pdfCell.setBorder(Border.NO_BORDER);

                    table.addCell(pdfCell);
                }
            } else {
                for (int j = 0; j < lastCellNum; j++) {
                    if(i == 11) {
                        HSSFCell cell = row.getCell(j);
                    }
                    HSSFCell cell = row.getCell(j);
                    if (cell != null) {
                        if (!isUse(cell)) {
                            CellRangeAddress cellRangeAddress = getCellRangeAddress(cell);
                            rowspan = 1;
                            colspan = 1;
                            maxWidth = 0;
                            if (cellRangeAddress != null) {
                                colspan = cellRangeAddress.getLastColumn() - cellRangeAddress.getFirstColumn() + 1;
                                rowspan = cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow() + 1;
                                j = cellRangeAddress.getLastColumn();
                                for(int k = cellRangeAddress.getFirstColumn() ; k < cellRangeAddress.getLastColumn() ; k++) {
                                    maxWidth += this.columnWidths[k];
                                }
                            } else {
                                maxWidth = this.columnWidths[j];
                            }
                            Cell pdfCell = transformCommon(cell, rowspan, colspan, maxWidth);
                            table.addCell(pdfCell);
                        }
                    } else {
                        // 补偿
                        Cell pdfCell = new Cell();
                        pdfCell.setBorder(Border.NO_BORDER);

                        table.addCell(pdfCell);
                    }
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
    private Cell transformCommon(HSSFCell cell, int rowspan, int colspan, float maxWidth) throws IOException {
        String value = Excel2PdfHelper.getValue(cell);

        Cell pdfCell = new Cell(rowspan, colspan)
                //.setHeight(cell.getRow().getHeight() * this.rate * 1.2f)
                .setHeight(cell.getRow().getHeightInPoints() * 1.2f)
                .setPadding(0);
        if (value.startsWith("${")) {
            pdfCell.setBorder(Border.NO_BORDER);
            annotationsCellMap.put(value, pdfCell);
        } else {
            Text text = new Text(value);
            setPdfCellFont(cell, text);
            HSSFCellStyle cellStyle = cell.getCellStyle();
            Paragraph paragraph = new Paragraph(text).setPadding(0f).setMargin(0f);
            if(cellStyle.getWrapText()) {
                paragraph.setMaxWidth(maxWidth);
            }

            pdfCell.add(paragraph);
            // 布局
            VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();
            pdfCell.setVerticalAlignment(Excel2PdfHelper.getVerticalAlignment(verticalAlignment));
            HorizontalAlignment alignment = cellStyle.getAlignment();
            pdfCell.setTextAlignment(Excel2PdfHelper.getTextAlignment(alignment, cell.getCellType()));

            //边框
            Excel2PdfHelper.transformBorder(cell, pdfCell);

            //背景色
            short colorIndex = cellStyle.getFillForegroundColor();
            HSSFColor color = this.customPalette.getColor(colorIndex);
            if (color.getIndex() != 64) {
                short[] triplet = color.getTriplet();
                int r = triplet[0] + 32;
                int g = triplet[1] + 90;
                int b = triplet[2] + 60;
                if(r > 255) {
                    r = 255;
                }
                if(g > 255) {
                    g = 255;
                }
                if(b > 255) {
                    b = 255;
                }
                DeviceRgb deviceRgb = new DeviceRgb(r, g, b);
                pdfCell.setBackgroundColor(deviceRgb);
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
    private void setPdfCellFont(HSSFCell cell, Text text) throws IOException {
        HSSFCellStyle cellStyle = cell.getCellStyle();
        //字体大小
        HSSFFont font = cellStyle.getFont(cell.getSheet().getWorkbook());
        //short fontHeight = font.getFontHeight();
        short fontHeight = font.getFontHeightInPoints();
        if(this.fontPath != null && !this.fontPath.equals("")) {
            text.setFont(PdfFontFactory.createFont(this.fontPath, PdfEncodings.IDENTITY_H));
        } else {
            text.setFont(PdfFontFactory.createFont(System.getProperty("user.dir") + "/doc/font/SimHei.TTF", PdfEncodings.IDENTITY_H));
        }

        //text.setFontSize(fontHeight * rate * 1.05f);
        text.setFontSize(fontHeight);

        //字体颜色
        HSSFColor hssfColor = font.getHSSFColor(cell.getSheet().getWorkbook());
        if (hssfColor != null && hssfColor.getIndex() != 64) {
            short[] triplet = hssfColor.getTriplet();
            text.setFontColor(new DeviceRgb(triplet[0], triplet[1], triplet[2]));
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

    private HSSFPicture getHSSFPicture(HSSFCell cell) {
        HSSFPatriarch patriarch = sheet.getDrawingPatriarch();
        if (patriarch != null) {
            List<HSSFShape> children = patriarch.getChildren();
            for (HSSFShape shape : children) {
                HSSFPicture hssfPicture = (HSSFPicture) shape;
                HSSFClientAnchor clientAnchor = hssfPicture.getClientAnchor();
                if (cell.getRowIndex() == clientAnchor.getRow1() && cell.getColumnIndex() == clientAnchor.getCol1()) {
                    return hssfPicture;
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
            //int columnWidth = this.sheet.getColumnWidth(i);
            //float realWidth = columnWidth * rate;
            float realWidth = this.sheet.getColumnWidthInPixels(i);
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
    private boolean isUse(HSSFCell cell) {
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
    private CellRangeAddress getCellRangeAddress(HSSFCell cell) {
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
