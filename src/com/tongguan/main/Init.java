package com.tongguan.main;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 初始化表
 * @author Administrator
 */

public class Init {
    Method myMethod;
    private int sheetNumbers = 0;
    private int coordinateRuler = 0;
    private String inFilePath = "";//引入文件位置
    private String outFilePath = "";//输出文件位置

    public Init() {
    }

    public Init(String inFilePath, String outFilePath){
        this.inFilePath = inFilePath;
        this.outFilePath = outFilePath;
        doInit();
    }

    /**
     * 开始初始化
     */
    private void doInit() {
        //创建方法对象
        myMethod = new Method();
        FileController fileController = new FileController();//创建文件控制器

        //"E:\\IDEAWorkSpace\\excel\\src\\com\\tongguan\\main\\test.xlsx"
        Workbook sWorkBook = fileController.getExcelFile(inFilePath);//获取相关工作簿

        sheetNumbers = sWorkBook.getNumberOfSheets();
        Sheet sSheet = sWorkBook.getSheetAt(sheetNumbers-1);
        String[] sName = new String[sSheet.getLastRowNum()];
        List<Cell> sList=  myMethod.getColumnWithCol(sSheet,0);
        for (Cell cell : sList){
            cell.setCellType(CellType.STRING);
            System.out.println(cell.getStringCellValue());
        }

        //初始化客户列表
        XSSFWorkbook rWorkBook = new XSSFWorkbook();
        ExcelStyleController styleController = new ExcelStyleController(rWorkBook);
        Sheet rSheet = rWorkBook.createSheet("应收");
        for (int i = 0;i < sList.size();i++){
            Row row = rSheet.createRow(i+1);
            Cell cell = row.createCell(0);
            cell.setCellType(CellType.STRING);
            if ( i==0 ){
                cell.setCellValue("客户名称");
            }else{
                cell.setCellValue(sList.get(i).getStringCellValue());
            }
        }


        //初始化表头
        int count = sheetNumbers;
        Row row0 = rSheet.createRow(0);
        Row row1 = rSheet.getRow(1);
        int temp = 1;
        for (int i = 0 ; i < count ; i++){
            if (i==0){
                for (int j = 0; j < 3; j++){
                    row0.createCell(temp).setCellValue(sWorkBook.getSheetName(i)+"年");
                    if (j==0){
                        row1.createCell(temp).setCellValue("期初余额");
                    }else if (j==1){
                        row1.createCell(temp).setCellValue("借方金额");
                    }else {
                        row1.createCell(temp).setCellValue("贷方金额");
                    }
                    temp++;
                }
            }
             else if (i == count-1){
                for (int j = 0; j < 3; j++){
                    row0.createCell(temp).setCellValue(sWorkBook.getSheetName(i)+"年");
                    if (j==0){
                        row1.createCell(temp).setCellValue("借方金额");
                    }else if (j==1){
                        row1.createCell(temp).setCellValue("贷方金额");
                    }else {
                        row1.createCell(temp).setCellValue("期末余额");
                    }
                    temp++;
                }
            }
            else{
                for (int j = 0; j < 2; j++){
                    row0.createCell(temp).setCellValue(sWorkBook.getSheetName(i)+"年");
                    if (j == 0){
                        row1.createCell(temp).setCellValue("借方金额");
                    }else {
                        row1.createCell(temp).setCellValue("贷方金额");
                    }
                    temp++;
                }
            }
        }



        //可封装成对应的方法
        Cell cell = row0.createCell(temp);
        XSSFCellStyle cellStyle = (XSSFCellStyle) styleController.alignCenterWithCenter();
        cell.setCellStyle(cellStyle);
        cell.setCellValue("帐龄");
        int temp2 = temp;
//        rSheet.addMergedRegion(new CellRangeAddress(0,0,temp,temp+5));
        row1.createCell(temp++).setCellValue("1年以内");
        row1.createCell(temp++).setCellValue("1-2年");
        row1.createCell(temp++).setCellValue("2-3年");
        switch (sheetNumbers){
            case 3:
                row1.createCell(temp++).setCellValue("3年以上");
                row1.createCell(temp).setCellValue("校验");
                rSheet.addMergedRegion(new CellRangeAddress(0,0,temp2,temp2+3));
                coordinateRuler = 9;
                break;
            case 4:
                row1.createCell(temp++).setCellValue("3-4年");
                row1.createCell(temp++).setCellValue("4年以上");
                row1.createCell(temp).setCellValue("校验");
                rSheet.addMergedRegion(new CellRangeAddress(0,0,temp2,temp2+4));
                coordinateRuler = 11;
                break;
            case 5:
                row1.createCell(temp++).setCellValue("3-4年");
                row1.createCell(temp++).setCellValue("4-5年");
                row1.createCell(temp++).setCellValue("5年以上");
                row1.createCell(temp).setCellValue("校验");
                rSheet.addMergedRegion(new CellRangeAddress(0,0,temp2,temp2+5));
                coordinateRuler = 13;
                break;
            default:
                break;

        }


        //获取对应的数据
        sList.remove(0);
        System.out.println(sList.size());
        int sheetId =0;
        List<DataBean> dataBeanList = new ArrayList<>();
        for (Sheet sheet :sWorkBook){
            for (Cell customerCell : sList){
                customerCell.setCellType(CellType.STRING);
                String customer = customerCell.getStringCellValue();
                if (sheetId == 0){
                    addDataBeanToList(sheet,customer,dataBeanList,"期初余额");
                }
//                else if(sheetId == sheetNumbers-1){
//                    addDataBeanToList(sheet,customer,dataBeanList,"期末余额");
//                }
                addDataBeanToList(sheet,customer,dataBeanList,"本年借方");
                addDataBeanToList(sheet,customer,dataBeanList,"本年贷方");
            }
            sheetId++;
        }


//        System.out.println(dataBeanList.size());
        //将对应的数据提取出来，获取到文件中
        for (DataBean dataBean: dataBeanList){
            System.out.println(dataBean.getYear()+"年,"+"客户："+dataBean.getCustomer()+"标题："+dataBean.getTitleName()+"数值："+dataBean.getValue());
            String title;
            switch (dataBean.getTitleName().trim()){
                case "本年借方":
                    title = "借方金额";
                    break;
                case "本年贷方":
                    title = "贷方金额";
                    break;
                default:
                    title = dataBean.getTitleName().trim();
                    break;
            }
            Cell test = myMethod.selectCellByCustomerTitleYear(rSheet,dataBean.getCustomer(),title,dataBean.getYear());
            String data = dataBean.getValue();
            test.setCellType(CellType.NUMERIC);
            test.setCellValue(Double.valueOf(data));
            test.setCellStyle(styleController.dataFormatWithMonetary2());
        }


//      计算相关的数值
        for (int i = 0; i<sList.size();i++){
            Row row = rSheet.getRow(2+i);
            switch (sheetNumbers){
                case 5:
                    accountByFiveYear(row);
                    break;
                case 4:
                    accountByFourYear(row);
                    break;
                case 3:
                    accountByThreeYear(row);
                    break;
                default:
                    System.out.println("分页数量不符合要求!");
                    break;
            }
        }

        //求有关列的和
        for (int i = 1 ;i < rSheet.getRow(1).getLastCellNum() ;i++){
            myMethod.getSumWithColumn(rSheet,i,2,sList.size()+2);
        }

        //总和校验

        Row row = rSheet.getRow(rSheet.getLastRowNum());
        double[] data = getDataToArray(row);
        Cell checkCell = myMethod.getCellWithRowAndCol(rSheet,rSheet.getLastRowNum()+2,1);
        checkCell.setCellValue(doCheck(data));


        rSheet.autoSizeColumn(0);
        try {
            fileController.saveExcleFile(rWorkBook,outFilePath);
        } catch (IOException e) {
//            e.printStackTrace();
            System.out.println("保存文件失败！错误信息为："+e);
        }
    }

    /**
     * 添加DataBean到List中
     * @param sheet 操作的分页
     * @param customer  客户名称
     * @param dataBeanList  添加对应的List对象
     * @param title 标题
     */
    private void addDataBeanToList(Sheet sheet, String customer,List<DataBean> dataBeanList,String title) {
        Cell itemCell = myMethod.selectCellByCustomerAndTitle(sheet,customer,title);
        if (itemCell != null){
            itemCell.setCellType(CellType.STRING);
            String item = itemCell.getStringCellValue();
            DataBean dataBean = new DataBean(customer,title,sheet.getSheetName(),item);
            dataBeanList.add(dataBean);
        }
    }

    /**
     * 只有五个分页的数据计算方法
     * @param row 哪一行
     */
    private void accountByFiveYear(Row row){
        double[] colValue = getDataToArray(row);
        double[] accountAge = new double[6];
        colValue[11] = getArraySum(colValue);
        row.createCell(coordinateRuler-1).setCellValue(colValue[11]);
        if (colValue[10] >= (colValue[11]+colValue[10]-colValue[9])){
            accountAge[0] = colValue[11];
        }else{
            accountAge[0] = colValue[9];
        }
        System.out.println("1年以内:"+accountAge[0]);
        row.createCell(coordinateRuler).setCellValue(accountAge[0]);
        if ((colValue[10]+colValue[8]) >= (colValue[11]+colValue[10]
                -colValue[9]+colValue[8]-colValue[7])){
            accountAge[1] = colValue[11]-accountAge[0];
        }else{
            accountAge[1] = colValue[7];
        }
        System.out.println("1-2年:"+accountAge[1]);
        row.createCell(coordinateRuler+1).setCellValue(accountAge[1]);
        if ((colValue[10]+colValue[8]+colValue[6])>=(colValue[11]
                +colValue[10]-colValue[9]+colValue[8]-colValue[7]+colValue[6]-colValue[5])){
            accountAge[2] = colValue[11]-accountAge[0]-accountAge[1];
        }else {
            accountAge[2] = colValue[5];
        }
        System.out.println("2-3年:"+accountAge[2]);
        row.createCell(coordinateRuler+2).setCellValue(accountAge[2]);
        if ((colValue[10]+colValue[8]+colValue[6]+colValue[4])>=
                (colValue[11]+colValue[10]-colValue[9]+colValue[8]
                        -colValue[7]+colValue[6]-colValue[5]+colValue[4]
                        -colValue[3])){
                accountAge[3] = colValue[11]-accountAge[0]-accountAge[1]-accountAge[2];
        }else {
            accountAge[3] = colValue[3];
        }
        System.out.println("3-4年:"+accountAge[3]);
        row.createCell(coordinateRuler+3).setCellValue(accountAge[3]);
        if ((colValue[10]+colValue[8]+colValue[6]+colValue[4]+colValue[2])>=
                (colValue[11]+colValue[10]-colValue[9]+colValue[8]
                        -colValue[7]+colValue[6]-colValue[5]+colValue[4]
                        -colValue[3]+colValue[2]-colValue[1])){
            accountAge[4] = colValue[11]-accountAge[0]-accountAge[1]-accountAge[2]-accountAge[3];
        }else{
            accountAge[4] = colValue[1];
        }
        System.out.println("4-5年:"+accountAge[4]);
        row.createCell(coordinateRuler+4).setCellValue(accountAge[4]);
        accountAge[5] = colValue[11]-accountAge[4]-accountAge[3]-accountAge[2]-accountAge[1]-accountAge[0];
        System.out.println("5年以上:"+accountAge[5]);
        row.createCell(coordinateRuler+5).setCellValue(accountAge[5]);
        row.createCell(coordinateRuler+6).setCellValue(doCheck(colValue));

    }

    /**
     * 分页只有四页的时候
     * @param row 哪一行
     */
    public void accountByFourYear(Row row){
        double[] colValue = getDataToArray(row);
        double[] accountAge = new double[6];

        colValue[9] = getArraySum(colValue);
        row.createCell(coordinateRuler-1).setCellValue(colValue[9]);
        if (colValue[8] >= (colValue[9]+colValue[8]-colValue[7])){
            accountAge[0] = colValue[9];
        }else{
            accountAge[0] = colValue[7];
        }
        System.out.println("1年以内:"+accountAge[0]);
        row.createCell(coordinateRuler).setCellValue(accountAge[0]);
        if ((colValue[8]+colValue[6]) >= (colValue[9]+colValue[8]
                -colValue[7]+colValue[6]-colValue[5])){
            accountAge[1] = colValue[9]-accountAge[0];
        }else{
            accountAge[1] = colValue[5];
        }
        System.out.println("1-2年:"+accountAge[1]);
        row.createCell(coordinateRuler+1).setCellValue(accountAge[1]);
        if ((colValue[8]+colValue[6]+colValue[4])>=(colValue[9]
                +colValue[8]-colValue[7]+colValue[6]-colValue[5]+colValue[4]-colValue[3])){
            accountAge[2] = colValue[9]-accountAge[0]-accountAge[1];
        }else {
            accountAge[2] = colValue[3];
        }
        System.out.println("2-3年:"+accountAge[2]);
        row.createCell(coordinateRuler+2).setCellValue(accountAge[2]);
        if ((colValue[8]+colValue[6]+colValue[4]+colValue[2])>=
                (colValue[9]+colValue[8]-colValue[7]+colValue[6]
                        -colValue[5]+colValue[4]-colValue[3]+colValue[2]
                        -colValue[1])){
            accountAge[3] = colValue[9]-accountAge[0]-accountAge[1]-accountAge[2];
        }else {
            accountAge[3] = colValue[1];
        }
        System.out.println("3-4年:"+accountAge[3]);
        row.createCell(coordinateRuler+3).setCellValue(accountAge[3]);

        accountAge[4] = colValue[9]-accountAge[3]-accountAge[2]-accountAge[1]-accountAge[0];
        System.out.println("4年以上:"+accountAge[4]);
        row.createCell(coordinateRuler+4).setCellValue(accountAge[4]);
        row.createCell(coordinateRuler+5).setCellValue(doCheck(colValue));
    }

    /**
     * 分页只有3个分页的时候
     * @param row 哪一行
     */
    public void accountByThreeYear(Row row){
        double[] colValue = getDataToArray(row);
        double[] accountAge = new double[5];

        colValue[7] = getArraySum(colValue);
        row.createCell(coordinateRuler-1).setCellValue(colValue[7]);
        if (colValue[6] >= (colValue[7]+colValue[6]-colValue[5])){
            accountAge[0] = colValue[7];
        }else{
            accountAge[0] = colValue[5];
        }
        System.out.println("1年以内:"+accountAge[0]);
        row.createCell(coordinateRuler).setCellValue(accountAge[0]);


        if ((colValue[6]+colValue[4]) >= (colValue[7]+colValue[6]
                -colValue[5]+colValue[4]-colValue[3])){
            accountAge[1] = colValue[7]-accountAge[0];
        }else{
            accountAge[1] = colValue[3];
        }
        System.out.println("1-2年:"+accountAge[1]);
        row.createCell(coordinateRuler+1).setCellValue(accountAge[1]);


        if ((colValue[6]+colValue[4]+colValue[2])>=(colValue[7]
                +colValue[6]-colValue[5]+colValue[4]-colValue[3]+colValue[2]-colValue[1])){
            accountAge[2] = colValue[7]-accountAge[0]-accountAge[1];
        }else {
            accountAge[2] = colValue[3];
        }
        System.out.println("2-3年:"+accountAge[2]);
        row.createCell(coordinateRuler+2).setCellValue(accountAge[2]);

        accountAge[3] = colValue[7]-accountAge[2]-accountAge[1]-accountAge[0];
        System.out.println("3年以上:"+accountAge[3]);
        row.createCell(coordinateRuler+3).setCellValue(accountAge[3]);

        row.createCell(coordinateRuler+4).setCellValue(doCheck(colValue));
    }

    /**
     * 将Cell中的数据添加到Array中
     * @param row 哪一行
     */
    private double[] getDataToArray(Row row) {
        double[] colValue = new double[coordinateRuler-1];
        for (int i = 1;i < coordinateRuler;i++){
            Cell cell = row.getCell(i);
            if (cell == null){
                cell = row.createCell(i);
            }
            cell.setCellType(CellType.STRING);
            colValue[i-1] = Double.valueOf(cell.getStringCellValue().isEmpty()?"0.00":cell.getStringCellValue());
        }
        return colValue;
    }

    private double getArraySum(double[] array){
        double sum = 0.0;
        for (int i = 0 ; i<array.length;i++){
            if (i == 0){
                sum = sum + array[i];
            }else{
                if (i%2 == 0){
                    sum = sum - array[i];
                }else {
                    sum = sum + array[i];
                }
            }
        }
        return sum;
    }

    private double doCheck(double[] array){
        double sum = 0.0;
        for (int i = 0 ; i<array.length;i++){
            if (i == 0){
                sum = sum + array[i];
            }else if (i == array.length-1){
                sum = sum - array[i];
            }else{
                if (i%2 == 0){
                    sum = sum - array[i];
                }else {
                    sum = sum + array[i];
                }
            }
        }
        return sum;
    }

    private Cell getCellByRow(Row row ,int col){
        Cell cell = row.getCell(col);
        if (cell==null){
            cell = row.createCell(col);
        }
        return cell;
    }
}
