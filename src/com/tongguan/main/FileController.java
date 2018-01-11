package com.tongguan.main;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * 文件控制
 * @author Administrator
 */
public class FileController {
//    private Workbook wb;

    public FileController() {
    }

    /**
     * 获取文件wb
     *
     * @return Workbook对象
     */
    public Workbook getExcelFile(String inFilePath){
        Workbook wb = null;
        try {
//            File file = new File("E:\\IDEAWorkSpace\\excel\\src\\com\\tongguan\\main\\test.xlsx");
//            System.out.println(file.exists());
            InputStream inp = new FileInputStream(inFilePath);
            wb = WorkbookFactory.create(inp);
        } catch (IOException e) {
            e.printStackTrace();
            System.out.print("读取文件失败,失败信息："+e);
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            System.out.print("读取文件失败,失败信息："+e);
        }
//        Sheet sheet = wb.getSheetAt(1);
//        sheet.getSheetName();
        return wb;
    }

    public void saveExcleFile(XSSFWorkbook xssfWorkbook, String outFilePath) throws IOException {
        Workbook wb = xssfWorkbook;
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(outFilePath);
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.print("输出文件失败,失败信息："+e);
        }finally {
            fileOut.close();
        }


    }
}
