package com.github.zhangchunsheng.excel2pdf.excel2003;

import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfString;
import com.itextpdf.kernel.pdf.annot.PdfFreeTextAnnotation;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.renderer.CellRenderer;
import com.itextpdf.layout.renderer.DrawContext;
import com.itextpdf.layout.renderer.IRenderer;
import com.itextpdf.layout.renderer.TableRenderer;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.util.Map;

/**
 * <pre>
 * Created by Chunsheng Zhang on 2020/8/17.
 * </pre>
 *
 * @author <a href="https://github.com/zhangchunsheng">Chunsheng Zhang</a>
 */
public class OverlappingAnnotationTableRenderer extends TableRenderer {
    private Map<String, Cell> annotationsCellMap;

    private PdfDocument pdfDocument;

    public OverlappingAnnotationTableRenderer(Table modelElement, Map<String, Cell> annotationsCellMap, PdfDocument pdfDocument) {
        super(modelElement);
        this.annotationsCellMap = annotationsCellMap;
        this.pdfDocument = pdfDocument;
    }

    @Override
    public void drawChildren(DrawContext drawContext) {
        super.drawChildren(drawContext);

        for (Map.Entry<String, Cell> entry : annotationsCellMap.entrySet()) {
            CellRenderer cellRenderer1 = rows.get(entry.getValue().getRow())[entry.getValue().getCol()];
            Rectangle rect1 = cellRenderer1.getOccupiedAreaBBox();

            this.doAnnotation(rect1.getX(), rect1.getY(), entry.getKey());
        }
    }

    @Override
    public IRenderer getNextRenderer() {
        return new OverlappingAnnotationTableRenderer((Table) modelElement, annotationsCellMap, pdfDocument);
    }

    private void doAnnotation(float x, float y, String value) {
        Rectangle rect = new Rectangle(x, y, 100, 20);
        PdfString pdfString = new PdfString(value);
        PdfFreeTextAnnotation ann = new PdfFreeTextAnnotation(rect, pdfString);

        // Setting title to the annotation
        ann.setTitle(new PdfString("Peter Zhang"));
        // <</BBox [0 0 36.26521 18.95174 ] /Filter /FlateDecode /FormType 1 /Length 141 /Resources 197 0 R /Subtype /Form /Type /XObject >>
        // ann.setDefaultAppearance(new PdfString("//Helvetica 12 Tf 0 g")); // PDFBox
        ann.setDefaultAppearance(new PdfString("//Arial 20 Tf 0 g"));
        this.pdfDocument.getLastPage().addAnnotation(ann);
    }
}