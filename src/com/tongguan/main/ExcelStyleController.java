package com.tongguan.main;

import com.microsoft.schemas.office.visio.x2012.main.CellType;
import org.apache.poi.ss.usermodel.*;

/**
 * 设置对应的排列方式
 */
public class ExcelStyleController {
    private Workbook workbook;
    private CellStyle cellStyle;
    private DataFormat format;
    /**
     * 构造方法
     * @param workbook 工作铺对象
     */
     public ExcelStyleController(Workbook workbook){
         this.workbook = workbook;
         this.format= workbook.createDataFormat();
    }

    /**
     * 排列为，水平居中，垂直靠下
     * @return
     */
    public CellStyle alignCenterWithBotton(){
        cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        return  cellStyle;
    }

    /**
     * 排列为，跨行水平居中，垂直靠下
     * @return
     */
    public CellStyle alignCenterSelectionWithBotton(){
        cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        return  cellStyle;
    }

    /**
     *排列为，水平填充，垂直居中
     * @return
     */
    public  CellStyle alignFillWithCenter(){
        cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.FILL);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    /**
     * 排列为，水平居中，垂直居中
     * @return
     */
    public CellStyle alignCenterWithCenter(){
        cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    /**
     * 排列为，水平常规，垂直居中
     * @return
     */
    public CellStyle alignGeneralWithCenter(){
        cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.GENERAL);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    /**
     * 排列为，自动换行
     * @return
     */
    public CellStyle alignJustifyWithJustify(){
        cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.JUSTIFY);
        cellStyle.setVerticalAlignment(VerticalAlignment.JUSTIFY);
        return cellStyle;
    }

    /**
     * 排列为，水平靠左，垂直靠顶部
     * @return
     */
    public CellStyle alignLeftWithTop(){
        cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        return cellStyle;
    }

    /**
     * 排列为，水平靠右，垂直靠顶部
     * @return
     */
    public CellStyle alignRightWithTop(){
        cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        return cellStyle;
    }

/**
 *下面的方法是设置数据格式
 */
    /**
     * 设置单元格数据为货币格式
     * @return
     */
    public CellStyle dataFormatWithMonetary (){
        cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format.getFormat("#,##0"));
        return cellStyle;
    }

    public CellStyle dataFormatWithMonetary2 (){
        cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format.getFormat("0.00"));
        cellStyle.setDataFormat(format.getFormat("#,##"));

        return cellStyle;
    }

    public CellStyle dataFormatWithMonetary2 (CellStyle cellStyle){
        cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format.getFormat("0.00"));
        cellStyle.setDataFormat(format.getFormat("#,##"));

        return cellStyle;
    }
}
