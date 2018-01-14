package com.tongguan.main;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * 初始化表
 *
 * @author Administrator
 */

public class Init {
    Method myMethod;
    private int sheetNumbers = 0;
    private int coordinateRuler = 0;
    private String inFilePath = "";//引入文件位置
    private String outFilePath = "";//输出文件位置
    private FileController fileController;
    ExcelStyleController styleController;
    private Workbook sWorkBook;
    private XSSFWorkbook rWorkBook;
    private Sheet sSheet;
    private XSSFSheet rSheet;
    private List<Cell> sList;
    private List<Cell> listData;
    private List<DataBean> dataBeanList;

    public Init() {
    }

    public Init(String inFilePath, String outFilePath) {
        this.inFilePath = inFilePath;
        this.outFilePath = outFilePath;
        doInit();
    }

    /**
     * 开始初始化
     */
    private void doInit() {
        //初始化必要对象
        myMethod = new Method();
        fileController = new FileController();//创建文件控制器
        //"E:\\IDEAWorkSpace\\excel\\src\\com\\tongguan\\main\\test.xlsx"
        sWorkBook = fileController.getExcelFile(inFilePath);//获取相关工作簿
        sheetNumbers = sWorkBook.getNumberOfSheets();
        rWorkBook = new XSSFWorkbook();
        styleController = new ExcelStyleController(rWorkBook);
        initCustomerList();         //初始化客户列表
        initHeader();               //初始化表头
        getData();                  //获取数据
        putData();                  //填写数据
        doOperation();              //进行相关的操作，计算
        rSheet.autoSizeColumn(0);   //宽度自适应

        //将数据类型转换成数字形式
        //typeStringToNumber(rSheet);


        try {
            fileController.saveExcleFile(rWorkBook, outFilePath);
        } catch (IOException e) {
//            e.printStackTrace();
            System.out.println("保存文件失败！错误信息为：" + e);
        }
    }

    private void initCustomerList() {
        //获取客户列表
        sSheet = sWorkBook.getSheetAt(sheetNumbers - 1);
        sList = myMethod.getColumnWithCol(sSheet, 0);
        for (Cell cell : sList) {
            cell.setCellType(CellType.STRING);
            System.out.println(cell.getStringCellValue());
        }
        //初始化客户列表
        rSheet = rWorkBook.createSheet("应收");
        for (int i = 0; i < sList.size(); i++) {
            Row row = rSheet.createRow(i + 5);
            Cell cell = row.createCell(0);
            cell.setCellType(CellType.STRING);
            if (i == 0) {
                cell.setCellValue("客户名称");
            } else {
                cell.setCellValue(sList.get(i).getStringCellValue());
            }
        }
    }

    private void initHeader(){
        //初始化表头
        int count = sheetNumbers;
        Row row0 = rSheet.createRow(4);
        Row row1 = rSheet.getRow(5);
        int temp = 1;
        for (int i = 0; i < count; i++) {
            if (i == 0) {
                for (int j = 0; j < 3; j++) {
                    row0.createCell(temp).setCellValue(sWorkBook.getSheetName(i) + "年");
                    if (j == 0) {
                        row1.createCell(temp).setCellValue("期初余额");
                    } else if (j == 1) {
                        row1.createCell(temp).setCellValue("借方金额");
                    } else {
                        row1.createCell(temp).setCellValue("贷方金额");
                    }
                    temp++;
                }
            } else if (i == count - 1) {
                for (int j = 0; j < 3; j++) {
                    row0.createCell(temp).setCellValue(sWorkBook.getSheetName(i) + "年");
                    if (j == 0) {
                        row1.createCell(temp).setCellValue("借方金额");
                    } else if (j == 1) {
                        row1.createCell(temp).setCellValue("贷方金额");
                    } else {
                        row1.createCell(temp).setCellValue("期末余额");
                    }
                    temp++;
                }
            } else {
                for (int j = 0; j < 2; j++) {
                    row0.createCell(temp).setCellValue(sWorkBook.getSheetName(i) + "年");
                    if (j == 0) {
                        row1.createCell(temp).setCellValue("借方金额");
                    } else {
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
        cell.setCellValue("账龄");
        int temp2 = temp;
//        rSheet.addMergedRegion(new CellRangeAddress(0,0,temp,temp+5));
        row1.createCell(temp++).setCellValue("1年以内");
        row1.createCell(temp++).setCellValue("1-2年");
        row1.createCell(temp++).setCellValue("2-3年");
        switch (sheetNumbers) {
            case 3:
                row1.createCell(temp++).setCellValue("3年以上");
                row1.createCell(temp).setCellValue("校验");
                rSheet.addMergedRegion(new CellRangeAddress(4, 4, temp2, temp2 + 3));
                coordinateRuler = 9;
                break;
            case 4:
                row1.createCell(temp++).setCellValue("3-4年");
                row1.createCell(temp++).setCellValue("4年以上");
                row1.createCell(temp).setCellValue("校验");
                rSheet.addMergedRegion(new CellRangeAddress(4, 4, temp2, temp2 + 4));
                coordinateRuler = 11;
                break;
            case 5:
                row1.createCell(temp++).setCellValue("3-4年");
                row1.createCell(temp++).setCellValue("4-5年");
                row1.createCell(temp++).setCellValue("5年以上");
                row1.createCell(temp).setCellValue("校验");
                rSheet.addMergedRegion(new CellRangeAddress(4, 4, temp2, temp2 + 5));
                coordinateRuler = 13;
                break;
            default:
                break;
        }
    }



    private void getData(){
        //获取对应的数据
        sList.remove(0);
        System.out.println(sList.size());
        int sheetId = 0;
        dataBeanList = new ArrayList<>();
        for (Sheet sheet : sWorkBook) {
            for (Cell customerCell : sList) {
                customerCell.setCellType(CellType.STRING);
                String customer = customerCell.getStringCellValue();
                if (sheetId == 0) {
                    addDataBeanToList(sheet, customer, dataBeanList, "期初余额");
                } else if (sheetId == sheetNumbers - 1) {
//                    addDataBeanToList(sheet,customer,dataBeanList,"期末余额");
                    listData = myMethod.getColumnWithCol(sheet, 6);
                }
                addDataBeanToList(sheet, customer, dataBeanList, "本年借方");
                addDataBeanToList(sheet, customer, dataBeanList, "本年贷方");
            }
            sheetId++;
        }
        listData.remove(0);
    }

    private void putData(){
        //将对应的数据提取出来，获取到文件中
        for (DataBean dataBean : dataBeanList) {

            String title;
            switch (dataBean.getTitleName().trim()) {
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
            Cell test = myMethod.selectCellByCustomerTitleYear(rSheet, dataBean.getCustomer(), title, dataBean.getYear());
            Double data = dataBean.getValue();
            BigDecimal b = new BigDecimal(data);
            data = b.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
            DecimalFormat df = new DecimalFormat("###0.00 ");
            df.setRoundingMode(RoundingMode.HALF_UP);
            //保留两位小数且不用科学计数法，并使用千分位
            String value = df.format(data);
            test.setCellType(CellType.STRING);
            test.setCellValue(value);
            test.setCellStyle(styleController.dataFormatWithMonetary2());
            System.out.println(dataBean.getYear() + "年," + "客户：" + dataBean.getCustomer() + "标题：" + dataBean.getTitleName() + "数值：" + value);
        }
    }

    private void doOperation(){
        //      计算相关的数值
        for (int i = 0; i < sList.size(); i++) {
            Row row = rSheet.getRow(6 + i);
            switch (sheetNumbers) {
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
        for (int i = 1; i < rSheet.getRow(6).getLastCellNum(); i++) {
            myMethod.getSumWithColumn(rSheet, i, 6, sList.size() + 6);
        }

//        总和校验

        Row row = rSheet.getRow(rSheet.getLastRowNum());
        double[] data = getDataToArray(row);
        Cell cellCheckValue = rSheet.getRow(rSheet.getLastRowNum()).getCell(coordinateRuler - 1);
        cellCheckValue.setCellType(CellType.STRING);
        double checkData = Double.valueOf(cellCheckValue.getStringCellValue());

        Cell checkCell = myMethod.getCellWithRowAndCol(rSheet, rSheet.getLastRowNum() + 2, 1);
        checkCell.setCellValue(doCheck(data, checkData));
    }

    /**
     * 添加DataBean到List中
     *
     * @param sheet        操作的分页
     * @param customer     客户名称
     * @param dataBeanList 添加对应的List对象
     * @param title        标题
     */
    private void addDataBeanToList(Sheet sheet, String customer, List<DataBean> dataBeanList, String title) {
        Cell itemCell = myMethod.selectCellByCustomerAndTitle(sheet, customer, title);
        if (itemCell != null) {
            itemCell.setCellType(CellType.STRING);
            String item = itemCell.getStringCellValue();
            DataBean dataBean = new DataBean(customer, title, sheet.getSheetName(), Double.valueOf(item));
            dataBeanList.add(dataBean);
        }
    }

    /**
     * 只有五个分页的数据计算方法
     *
     * @param row 哪一行
     */
    private void accountByFiveYear(Row row) {
        double[] colValue = getDataToArray(row);
        double[] accountAge = new double[6];
        colValue[11] = getArraySum(colValue);
        row.createCell(coordinateRuler - 1).setCellValue(colValue[11]);
        if (colValue[10] >= (colValue[11] + colValue[10] - colValue[9])) {
            accountAge[0] = colValue[11];
        } else {
            accountAge[0] = colValue[9];
        }
        System.out.println("1年以内:" + accountAge[0]);
        row.createCell(coordinateRuler).setCellValue(changeType(accountAge[0]));
        if ((colValue[10] + colValue[8]) >= (colValue[11] + colValue[10]
                - colValue[9] + colValue[8] - colValue[7])) {
            accountAge[1] = colValue[11] - accountAge[0];
        } else {
            accountAge[1] = colValue[7];
        }
        System.out.println("1-2年:" + accountAge[1]);
        row.createCell(coordinateRuler + 1).setCellValue(changeType(accountAge[1]));
        if ((colValue[10] + colValue[8] + colValue[6]) >= (colValue[11]
                + colValue[10] - colValue[9] + colValue[8] - colValue[7] + colValue[6] - colValue[5])) {
            accountAge[2] = colValue[11] - accountAge[0] - accountAge[1];
        } else {
            accountAge[2] = colValue[5];
        }
        System.out.println("2-3年:" + accountAge[2]);
        row.createCell(coordinateRuler + 2).setCellValue(changeType(accountAge[2]));
        if ((colValue[10] + colValue[8] + colValue[6] + colValue[4]) >=
                (colValue[11] + colValue[10] - colValue[9] + colValue[8]
                        - colValue[7] + colValue[6] - colValue[5] + colValue[4]
                        - colValue[3])) {
            accountAge[3] = colValue[11] - accountAge[0] - accountAge[1] - accountAge[2];
        } else {
            accountAge[3] = colValue[3];
        }
        System.out.println("3-4年:" + accountAge[3]);
        row.createCell(coordinateRuler + 3).setCellValue((accountAge[3]));
        if ((colValue[10] + colValue[8] + colValue[6] + colValue[4] + colValue[2]) >=
                (colValue[11] + colValue[10] - colValue[9] + colValue[8]
                        - colValue[7] + colValue[6] - colValue[5] + colValue[4]
                        - colValue[3] + colValue[2] - colValue[1])) {
            accountAge[4] = colValue[11] - accountAge[0] - accountAge[1] - accountAge[2] - accountAge[3];
        } else {
            accountAge[4] = colValue[1];
        }
        System.out.println("4-5年:" + accountAge[4]);
        row.createCell(coordinateRuler + 4).setCellValue(changeType(accountAge[4]));
        accountAge[5] = colValue[11] - accountAge[4] - accountAge[3] - accountAge[2] - accountAge[1] - accountAge[0];
        System.out.println("5年以上:" + accountAge[5]);
        row.createCell(coordinateRuler + 5).setCellValue(changeType(accountAge[5]));
        Cell checkCell = listData.get(row.getRowNum() - 6);
//        System.out.println(Double.valueOf(changeType(Double.valueOf(checkCell.getStringCellValue()))));
        row.createCell(coordinateRuler + 6).setCellValue(doCheck(colValue,Double.valueOf(checkCell.getStringCellValue())));

    }

    /**
     * 分页只有四页的时候
     *
     * @param row 哪一行
     */
    public void accountByFourYear(Row row) {
        double[] colValue = getDataToArray(row);
        double[] accountAge = new double[6];

        colValue[9] = getArraySum(colValue);
        row.createCell(coordinateRuler - 1).setCellValue(changeType(colValue[9]));
        if (colValue[8] >= (colValue[9] + colValue[8] - colValue[7])) {
            accountAge[0] = colValue[9];
        } else {
            accountAge[0] = colValue[7];
        }
        System.out.println("1年以内:" + accountAge[0]);
        row.createCell(coordinateRuler).setCellValue(changeType(accountAge[0]));
        if ((colValue[8] + colValue[6]) >= (colValue[9] + colValue[8]
                - colValue[7] + colValue[6] - colValue[5])) {
            accountAge[1] = colValue[9] - accountAge[0];
        } else {
            accountAge[1] = colValue[5];
        }
        System.out.println("1-2年:" + accountAge[1]);
        row.createCell(coordinateRuler + 1).setCellValue(changeType(accountAge[1]));
        if ((colValue[8] + colValue[6] + colValue[4]) >= (colValue[9]
                + colValue[8] - colValue[7] + colValue[6] - colValue[5] + colValue[4] - colValue[3])) {
            accountAge[2] = colValue[9] - accountAge[0] - accountAge[1];
        } else {
            accountAge[2] = colValue[3];
        }
        System.out.println("2-3年:" + accountAge[2]);
        row.createCell(coordinateRuler + 2).setCellValue(changeType(accountAge[2]));
        if ((colValue[8] + colValue[6] + colValue[4] + colValue[2]) >=
                (colValue[9] + colValue[8] - colValue[7] + colValue[6]
                        - colValue[5] + colValue[4] - colValue[3] + colValue[2]
                        - colValue[1])) {
            accountAge[3] = colValue[9] - accountAge[0] - accountAge[1] - accountAge[2];
        } else {
            accountAge[3] = colValue[1];
        }
        System.out.println("3-4年:" + accountAge[3]);
        row.createCell(coordinateRuler + 3).setCellValue(changeType(accountAge[3]));

        accountAge[4] = colValue[9] - accountAge[3] - accountAge[2] - accountAge[1] - accountAge[0];
        System.out.println("4年以上:" + accountAge[4]);
        row.createCell(coordinateRuler + 4).setCellValue(changeType(accountAge[4]));
        Cell checkCell = listData.get(row.getRowNum() - 6);
        row.createCell(coordinateRuler + 5).setCellValue(doCheck(colValue, Double.valueOf(checkCell.getStringCellValue())));
    }

    /**
     * 分页只有3个分页的时候
     *
     * @param row 哪一行
     */
    public void accountByThreeYear(Row row) {
        double[] colValue = getDataToArray(row);
        double[] accountAge = new double[5];

        colValue[7] = getArraySum(colValue);
        row.createCell(coordinateRuler - 1).setCellValue(changeType(colValue[7]));
        if (colValue[6] >= (colValue[7] + colValue[6] - colValue[5])) {
            accountAge[0] = colValue[7];
        } else {
            accountAge[0] = colValue[5];
        }
        System.out.println("1年以内:" + accountAge[0]);
        row.createCell(coordinateRuler).setCellValue(changeType(accountAge[0]));


        if ((colValue[6] + colValue[4]) >= (colValue[7] + colValue[6]
                - colValue[5] + colValue[4] - colValue[3])) {
            accountAge[1] = colValue[7] - accountAge[0];
        } else {
            accountAge[1] = colValue[3];
        }
        System.out.println("1-2年:" + accountAge[1]);
        row.createCell(coordinateRuler + 1).setCellValue(changeType(accountAge[1]));


        if ((colValue[6] + colValue[4] + colValue[2]) >= (colValue[7]
                + colValue[6] - colValue[5] + colValue[4] - colValue[3] + colValue[2] - colValue[1])) {
            accountAge[2] = colValue[7] - accountAge[0] - accountAge[1];
        } else {
            accountAge[2] = colValue[3];
        }
        System.out.println("2-3年:" + accountAge[2]);
        row.createCell(coordinateRuler + 2).setCellValue(changeType(accountAge[2]));

        accountAge[3] = colValue[7] - accountAge[2] - accountAge[1] - accountAge[0];
        System.out.println("3年以上:" + accountAge[3]);
        row.createCell(coordinateRuler + 3).setCellValue(changeType(accountAge[3]));

        int test = row.getRowNum();
        Cell checkCell = listData.get(test - 6);
        double test1 =  Double.valueOf(changeType(Double.valueOf(checkCell.getStringCellValue())));
        double test2 = doCheck(colValue,test1);
        row.createCell(coordinateRuler + 4).setCellValue(test2);
    }

    /**
     * 将Cell中的数据添加到Array中
     *
     * @param row 哪一行
     */
    private double[] getDataToArray(Row row) {
        double[] colValue = new double[coordinateRuler - 1];
        for (int i = 1; i < coordinateRuler; i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                cell = row.createCell(i);
            }
            cell.setCellType(CellType.STRING);
            String data = cell.getStringCellValue();
            if (data.isEmpty()){
                colValue[i - 1] = 0.00;
            }else{
                data = changeType(Double.valueOf(data));
                colValue[i - 1] = Double.valueOf(data);
            }
        }
        return colValue;
    }

    /**
     * 获取数组算法计算期末余额
     * @param array 需要计算的数组
     * @return 返回期末余额
     */
    private double getArraySum(double[] array) {
        double sum = 0.0;
        for (int i = 0; i < array.length; i++) {
            if (i == 0) {
                sum = sum + array[i];
            } else {
                if (i % 2 == 0) {
                    sum = sum - array[i];
                } else {
                    sum = sum + array[i];
                }
            }
        }
        sum = Double.valueOf(changeType(sum));
        return sum;
    }

    /**
     *  校验期末余额是否正确
     * @param array
     * @param value
     * @return
     */
    private double doCheck(double[] array, double value) {
        double check = 0.0;
        for (int i = 0; i < array.length - 1; i++) {
            if (i == 0) {
                check = check + array[i];
            } else if (i == array.length - 2) {
                check = check - array[i];
            } else {
                if (i % 2 == 0) {
                    check = check - array[i];
                } else {
                    check = check + array[i];
                }
            }
        }
        check = check - value;
        check = Double.valueOf(changeType(check));
        return check;
    }

    /**
     * 改变数据格式为小数点后两位
     * @param data
     * @return
     */
    private String changeType(double data){
        BigDecimal b = new BigDecimal(data);
        data = b.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
        DecimalFormat df = new DecimalFormat("###0.00 ");
        df.setRoundingMode(RoundingMode.HALF_UP);
        //保留两位小数且不用科学计数法，并不使用千分位
        String value = df.format(data);
        return value;
    }

    private void typeStringToNumber(Sheet sheet){
        for (Row row:sheet){
            for (Cell cell:row){
                cell.setCellType(CellType.NUMERIC);
            }
        }
    }

}
