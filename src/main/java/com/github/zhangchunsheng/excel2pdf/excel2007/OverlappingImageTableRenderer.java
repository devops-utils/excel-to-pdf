package com.github.zhangchunsheng.excel2pdf.excel2007;

import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.renderer.CellRenderer;
import com.itextpdf.layout.renderer.DrawContext;
import com.itextpdf.layout.renderer.IRenderer;
import com.itextpdf.layout.renderer.TableRenderer;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFSheet;

/**
 * @date 2021/8/17
 */
public class OverlappingImageTableRenderer extends TableRenderer {
    private HSSFPicture picture;

    private HSSFSheet sheet;

    public OverlappingImageTableRenderer(Table modelElement, HSSFPicture picture, HSSFSheet sheet) {
        super(modelElement);
        this.picture = picture;
        this.sheet = sheet;
    }

    @Override
    public void drawChildren(DrawContext drawContext) {
        super.drawChildren(drawContext);

        HSSFClientAnchor clientAnchor = picture.getClientAnchor();
        // Use the coordinates of the cell in the fourth row and the second column to draw the image
        CellRenderer cellRenderer1 = rows.get(clientAnchor.getRow1())[clientAnchor.getCol1()];
        Rectangle rect1 = cellRenderer1.getOccupiedAreaBBox();
        CellRenderer cellRenderer2 = rows.get(clientAnchor.getRow2())[clientAnchor.getCol2()];
        Rectangle rect2 = cellRenderer2.getOccupiedAreaBBox();

        float widthRate = (super.getOccupiedAreaBBox().getWidth() + rect2.getWidth()) / getExcelWidth(sheet);
        float heightRate = (super.getOccupiedAreaBBox().getHeight() - rect2.getHeight()) / getExcelHeight(sheet);


//        float imgX1 = rect1.getLeft() + clientAnchor.getDx1() * widthRate;
//        float imgX2 = rect2.getLeft() + clientAnchor.getDx2() * widthRate;
//        float imgY1 = rect1.getTop() - clientAnchor.getDy1() * heightRate;
//        float imgY2 = rect2.getTop() - clientAnchor.getDy2() * heightRate;

//        float height = Math.abs(imgY2 - imgY1);
//        float width = Math.abs(imgX2 - imgX1);

        float width = 0f;
        for (int i = clientAnchor.getCol1(); i < clientAnchor.getCol2(); i++) {
            width += sheet.getColumnWidth(i);
        }
        width = Math.abs(width - clientAnchor.getDx1() + clientAnchor.getDx2()) * widthRate;

        float height = 0f;
        for (int i = clientAnchor.getRow1(); i < clientAnchor.getRow2(); i++) {
            height += sheet.getRow(i).getHeight();
        }
        height = Math.abs(height - clientAnchor.getDy1() + clientAnchor.getDy2()) * heightRate;

        float x = rect1.getLeft() + clientAnchor.getDx1() * widthRate;
        float y = rect1.getTop() - height - clientAnchor.getDy1() * heightRate;
        ImageData imageData = ImageDataFactory.create(picture.getPictureData().getData());
        drawContext.getCanvas().addImage(imageData, width, 0, 0, height, x, y);
    }

    private float getExcelHeight(HSSFSheet sheet) {
        int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
        float result = 0;
        for (int i = 0; i < physicalNumberOfRows; i++) {
            result += sheet.getRow(i).getHeight();
        }
        return result;
    }

    private float getExcelWidth(HSSFSheet sheet) {
        short lastCellNum = sheet.getRow(0).getLastCellNum();
        float result = 0;
        for (int i = 0; i < lastCellNum; i++) {
            result += sheet.getColumnWidth(i);
        }
        return result;
    }

    @Override
    public IRenderer getNextRenderer() {
        return new OverlappingImageTableRenderer((Table) modelElement, picture, sheet);
    }
}