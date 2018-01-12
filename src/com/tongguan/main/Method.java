package com.tongguan.main;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import sun.misc.Cleaner;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * 操作表的一些常规方法
 * @author tianjun
 */
public class Method {
    public Method(){

    }

    /**
     * 通过输入行列坐标，获取Cell对象
     *
     * @param sheet 表分页
     * @param row 行
     * @param col 列
     * @return Cell对象
     */
    public Cell getCellWithRowAndCol(Sheet sheet, int row, int col){
        Cell cell = null;
        Row sheetRow = sheet.getRow(row);
        if (sheetRow == null){
            sheetRow = sheet.createRow(row);
        }
        cell = sheetRow.getCell(col);
        if (cell == null){
            cell = sheetRow.createCell(col);
        }
        return cell;
    }




    /**
     * 获取一列的Cell，通过List返回
     *
     * @param sheet  表分页
     * @param col   列
     * @return List 对象结果
     */
    public List<Cell> getColumnWithCol(Sheet sheet,int col){
        List<Cell> sheetCol = new ArrayList<>();
        for (Row row:sheet){
            Cell cell = row.getCell(col);
            cell.setCellType(CellType.STRING);
            sheetCol.add(row.getCell(col));
        }

        return sheetCol;
    }

    /**
     * 通过值获取List中的索引，及获取行的值
     *
     * @param sheetCol Cell泛型的List对象
     * @param value 需要查找索引的值
     * @return 返回对应的索引，当为-1时，遍历失败，未找到对应的值
     */
    public int getListIdByValue(List<Cell> sheetCol, String value){
        int id = -1;
        for (Cell cell: sheetCol){
            cell.setCellType(CellType.STRING);
            if (cell.getStringCellValue().equals(value)){
                id = sheetCol.indexOf(cell);
            }
        }
        return id;
    }

    /**
     * 通过值获取Row中的索引，及获取行中的列的索引
     *
     * @param row 行对象
     * @param value 需要查找的值的字符串
     * @return 返回对应的索引，当为-1时，遍历失败，未找到对应的值
     */
    public int getColByRowValue(Row row,String value){
        int id = -1;
        for (Cell cell: row){
            cell.setCellType(CellType.STRING);
            id++;
            if (cell.getStringCellValue().equals(value)){
                return id;
            }
        }
        return -1;
    }

    /**
     * 获取相同值的第一个Cell
     * @param sheet 表个分页
     * @param value 查找的值
     * @return 返回对应的Cell
     */
    public Cell getCellByValue(Sheet sheet, String value){
        int[] index = new int[2];
        for (Row row:sheet){
            for (Cell cell:row){
                cell.setCellType(CellType.STRING);
                if (cell.getStringCellValue().equals(value)){
                    return cell;
                }
            }
        }
        return null;
    }


    /**
     * 通过两个Cell获取到相应的行列的Cell
     * 获取客户的行，获取标题的列
     * @param sheet 获取哪个分页
     * @param customer 获取客户的行
     * @param title 获取标题的列
     * @return 返回对应的cell对象
     */
    public Cell selectCellByRcellAndCcell(Sheet sheet,Cell customer,Cell title){
        int row = customer.getRowIndex();
        int col = title.getColumnIndex();
        Row row1 = sheet.getRow(row);
        Cell cell = row1.getCell(col);
        return cell;
    }

    /**
     * 通过客户姓名和标题获取对应的Cell
     * @param sheet 获取哪个分页
     * @param customer 客户的名字String
     * @param title 表单标题String
     * @return 返回对应的Cell 只要有一个为空，就查询失败
     */
    public Cell selectCellByCustomerAndTitle(Sheet sheet,String customer,String title){
        Cell cCell = getCellByValue(sheet,customer);
        if (cCell == null){
            return null;
        }
        Cell tCell = getCellByValue(sheet,title);
        if (tCell == null){
            return null;
        }
        return selectCellByRcellAndCcell(sheet,cCell,tCell);
    }

    /**
     * 查找所有有关的值的cell并保存在List中
     * @param sheet 获取哪个分页，操作对象
     * @param value 对比哪个值
     * @return
     */
    public List<Cell> selectCellsByValue(Sheet sheet,String value){
        List<Cell> cells = new ArrayList<>();
        for (Row row:sheet){
            for (Cell cell:row){
                cell.setCellType(CellType.STRING);
                if (cell.getStringCellValue().equals(value)){
                    cells.add(cell);
                }
            }
        }
        return cells;
    }

    /**
     * 将使用三个属性对数据进行定位
     * @param sheet 操作是分页
     * @param customer 顾客
     * @param title 标题
     * @param year 年份
     * @return 返回对应位置的Cell
     */
    public Cell selectCellByCustomerTitleYear(Sheet sheet,String customer,String title,String year){
        List<Cell> years = selectCellsByValue(sheet,year+"年");
        Row row = sheet.getRow(1);
        int iCol = 0;
        for (Cell cell:years){
            iCol = cell.getColumnIndex();
            String string = row.getCell(iCol).getStringCellValue();
            if (string.equals(title)){
                break;
            }
        }
        Cell cell = getCellByValue(sheet,customer);
        int iRow = cell.getRowIndex();
        return sheet.getRow(iRow).createCell(iCol);
    }

    /**
     * 对列求和,并添加到最后一列
     * @param sheet 哪一张分页
     * @param col   列的位置
     */
    public double getSumWithColumn(Sheet sheet,int col,int starRow,int endRow){
        double sum = 0.00;
        for (int i = starRow;i<=endRow;i++){
            Row row = sheet.getRow(i);
            if (row == null){
                row = sheet.createRow(i);
            }
            Cell cell = row.getCell(col);
            if (cell == null){
                cell = row.createCell(col);
                cell.setCellValue(0.0);
            }else {
                cell.setCellType(CellType.STRING);
                String stringCellValue = cell.getStringCellValue();
                if (stringCellValue.isEmpty()){
                    stringCellValue = "0.00";
                }
                sum = sum + Double.valueOf(stringCellValue);
            }
        }
        BigDecimal b = new BigDecimal(sum);
        sum = b.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
        DecimalFormat df = new DecimalFormat("###0.00 ");
        df.setRoundingMode(RoundingMode.HALF_UP);
        //保留两位小数且不用科学计数法，并不使用千分位
        String value = df.format(sum);

        int rowIndex = sheet.getLastRowNum();
        Row row = sheet.getRow(rowIndex);
        if (row == null){
            row = sheet.createRow(rowIndex);
        }
        Cell sumCell =  row.createCell(col);
        sumCell.setCellValue(value);
        return sum;
    }

    public void createContentType(Workbook workbook){
        CellStyle cellStyle = workbook.createCellStyle();
        for (Sheet sheet:workbook){
            for (Row row:sheet){
                for (Cell cell: row){
                    cell.setCellType(CellType.STRING);
                    String data = cell.getStringCellValue();

                    Boolean isNum = false;//data是否为数值型
                    Boolean isInteger=false;//data是否为整数
                    Boolean isPercent=false;//data是否为百分数
                    Boolean isScience = false;//data是否是科学计数法
                    if (data != null || "".equals(data)) {
                        //判断data是否为数值型
                        isNum = data.toString().matches("^(-?\\d+)(\\.\\d+)?$");
                        //判断data是否为整数（小数部分是否为0）
                        isInteger=data.toString().matches("^[-\\+]?[\\d]*$");
                        //判断data是否为百分数（是否包含“%”）
                        isPercent=data.toString().contains("%");
                        isScience = data.matches("^((-?\\\\d+.?\\\\d*)[Ee]{1}(-?\\\\d+))$");
                    }
                    if (isNum && !isPercent) {
                        XSSFDataFormat df = (XSSFDataFormat) workbook.createDataFormat(); // 此处设置数据格式
                        if (isInteger) {
                            cellStyle.setDataFormat(df.getFormat("#,#0"));//数据格式只显示整数
                        }else{
                            cellStyle.setDataFormat(df.getFormat("#,##0.00"));//保留两位小数点
                        }
                        // 设置单元格格式
                        cell.setCellStyle(cellStyle);
                        // 设置单元格内容为double类型
                        cell.setCellValue(Double.parseDouble(data.toString()));
                    } else {
                        cell.setCellStyle(cellStyle);
                        // 设置单元格内容为字符型
                        cell.setCellValue(data.toString());
                    }
                }
            }
        }
    }
}
