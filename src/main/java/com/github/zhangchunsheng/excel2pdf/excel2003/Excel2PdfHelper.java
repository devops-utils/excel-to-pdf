package com.github.zhangchunsheng.excel2pdf.excel2003;

import com.itextpdf.kernel.colors.Color;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.layout.borders.*;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.property.HorizontalAlignment;
import com.itextpdf.layout.property.TextAlignment;
import com.itextpdf.layout.property.VerticalAlignment;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.util.LocaleUtil;

import java.text.SimpleDateFormat;

/**
 * <pre>
 * Created by Chunsheng Zhang on 2020/8/17.
 * </pre>
 *
 * @author <a href="https://github.com/zhangchunsheng">Chunsheng Zhang</a>
 */
public class Excel2PdfHelper {

    public static String getValue(HSSFCell cell) {
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case BOOLEAN:
                return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy", LocaleUtil.getUserLocale());
                    sdf.setTimeZone(LocaleUtil.getUserTimeZone());
                    return sdf.format(cell.getDateCellValue());
                }
                return String.valueOf(cell.getNumericCellValue());
            case STRING:
                return cell.getStringCellValue();
            default:
                return "";
        }
    }

    public static VerticalAlignment getVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment verticalAlignment) {
        switch (verticalAlignment) {
            case TOP:
                return VerticalAlignment.TOP;
            case BOTTOM:
                return VerticalAlignment.BOTTOM;
            case JUSTIFY:
            case CENTER:
                return VerticalAlignment.MIDDLE;
        }
        return VerticalAlignment.MIDDLE;
    }

    public static HorizontalAlignment getHorizontalAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment alignment) {
        switch (alignment) {
            case LEFT:
                return HorizontalAlignment.LEFT;
            case RIGHT:
                return HorizontalAlignment.RIGHT;
            case CENTER:
            case FILL:
            case GENERAL:
                return HorizontalAlignment.CENTER;
        }
        return HorizontalAlignment.CENTER;
    }

    public static TextAlignment getTextAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment alignment, CellType cellType) {
        switch (alignment) {
            case LEFT:
                return TextAlignment.LEFT;
            case RIGHT:
                return TextAlignment.RIGHT;
            case CENTER:
                return TextAlignment.CENTER;
            case JUSTIFY:
                return TextAlignment.JUSTIFIED;
            case GENERAL:
                if (cellType == CellType.NUMERIC) {
                    return TextAlignment.RIGHT;
                } else if (cellType == CellType.BOOLEAN) {
                    return TextAlignment.CENTER;
                }
        }
        return TextAlignment.LEFT;
    }

    /**
     * 处理边框
     *
     * @param cell
     * @param pdfCell
     */
    public static void transformBorder(HSSFCell cell, Cell pdfCell) {
        HSSFCellStyle cellStyle = cell.getCellStyle();
        BorderStyle borderBottom = cellStyle.getBorderBottom();
        pdfCell.setBorderBottom(getBorder(borderBottom, cellStyle.getBottomBorderColor(), cell));

        BorderStyle borderLeft = cellStyle.getBorderLeft();
        pdfCell.setBorderLeft(getBorder(borderLeft, cellStyle.getLeftBorderColor(), cell));

        BorderStyle borderRight = cellStyle.getBorderRight();
        pdfCell.setBorderRight(getBorder(borderRight, cellStyle.getRightBorderColor(), cell));

        BorderStyle borderTop = cellStyle.getBorderTop();
        pdfCell.setBorderTop(getBorder(borderTop, cellStyle.getTopBorderColor(), cell));
    }

    public static Border getBorder(BorderStyle borderStyle, short colorIndex, HSSFCell cell) {
        HSSFPalette customPalette = cell.getSheet().getWorkbook().getCustomPalette();
        HSSFColor color = customPalette.getColor(colorIndex);
        Color defaultColor = ColorConstants.BLACK;
        if (color != null && color.getIndex() != 64) {
            short[] triplet = color.getTriplet();
            defaultColor = new DeviceRgb(triplet[0], triplet[1], triplet[2]);
        }
        switch (borderStyle) {
            case THIN:
                return new SolidBorder(defaultColor, 0.3f);
            case MEDIUM:
                return new SolidBorder(defaultColor, 0.5f);
            case DASHED:
                return new DashedBorder(defaultColor, 0.3f);
            case DOTTED:
                return new DottedBorder(defaultColor, 0.3f);
            case THICK:
                return new SolidBorder(defaultColor, 1f);
            case DOUBLE:
                return new DoubleBorder(defaultColor, 0.3f);
            case MEDIUM_DASHED:
                return new DashedBorder(defaultColor, 0.5f);
        }
        return Border.NO_BORDER;
    }
}
