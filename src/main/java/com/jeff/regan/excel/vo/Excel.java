package com.jeff.regan.excel.vo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * 建立编辑excel工具
 *
 * @author zhangby
 * @date 2017/8/2 15:08
 */
public class Excel {
    private Workbook workbook;

    /**
     * 创建excel
     */
    public Excel() {
        //创建工作簿对象
        this.workbook = new HSSFWorkbook();
    }

    /**
     * 读取excel模板,创建excel
     *
     * @param filePath 模板全路径
     * @throws IOException
     */
    public Excel(String filePath) throws IOException {
        //excel模板路径
        File fi = new File(filePath);
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(fi));
        //读取excel模板
        try {
            this.workbook = new HSSFWorkbook(fs);
        }catch (Exception e){
            this.workbook = new XSSFWorkbook(new FileInputStream(fi));
        }
    }

    /**
     * 创建工作簿
     *
     * @return
     */
    public ExcelSheet createSheet() {
        return new ExcelSheet(this.workbook.createSheet());
    }

    /**
     * 根据工作簿名称创建工作簿
     *
     * @param sheetName 工作簿名称
     * @return
     */
    public ExcelSheet createSheet(String sheetName) {
        return new ExcelSheet(this.workbook.createSheet(sheetName));
    }

    /**
     * 根据工作簿 index 创建工作簿
     *
     * @param sheetIndex
     * @return
     */
    public ExcelSheet createSheet(int sheetIndex) {
        return new ExcelSheet(this.workbook.cloneSheet(sheetIndex));
    }

    /**
     * 获取工作簿 根据sheet name
     * @param sheetName
     * @return
     */
    public ExcelSheet getSheet(String sheetName) {
        return new ExcelSheet(this.workbook.getSheet(sheetName));
    }

    /**
     * 获取工作簿 根据
     * @param sheetIndex
     * @return
     */
    public ExcelSheet getSheet(int sheetIndex) {
        return new ExcelSheet(this.workbook.getSheetAt(sheetIndex));
    }
    /**
     * 获取工作簿 根据
     * @return
     */
    public ExcelSheet getSheet() {
        return new ExcelSheet(this.workbook.getSheetAt(0));
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * 根据路径到处excel
     *
     * @param file
     * @throws IOException
     */
    public void saveExcel(String file) throws IOException {
        OutputStream out = new FileOutputStream(new File(file));
        this.workbook.write(out);
    }

    /**
     * excel导出
     *
     * @param outputStream
     * @throws IOException
     */
    public void saveExcel(OutputStream outputStream) throws IOException {
        this.workbook.write(outputStream);
    }
}
