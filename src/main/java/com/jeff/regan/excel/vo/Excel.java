package com.jeff.regan.excel.vo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;

/**
 * 建立编辑excel工具
 *
 * @author zhangby
 * @date 2017/8/2 15:08
 */
public class Excel {
    private HSSFWorkbook hssfWorkbook;

    /**
     * 创建excel
     */
    public Excel() {
        //创建工作簿对象
        this.hssfWorkbook = new HSSFWorkbook();
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
        this.hssfWorkbook = new HSSFWorkbook(fs);
    }

    /**
     * 创建工作簿
     *
     * @return
     */
    public ExcelSheet createSheet() {
        return new ExcelSheet(this.hssfWorkbook.createSheet());
    }

    /**
     * 根据工作簿名称创建工作簿
     *
     * @param sheetName 工作簿名称
     * @return
     */
    public ExcelSheet createSheet(String sheetName) {
        return new ExcelSheet(this.hssfWorkbook.createSheet(sheetName));
    }

    /**
     * 根据工作簿 index 创建工作簿
     *
     * @param sheetIndex
     * @return
     */
    public ExcelSheet createSheet(int sheetIndex) {
        return new ExcelSheet(this.hssfWorkbook.cloneSheet(sheetIndex));
    }

    /**
     * 获取工作簿 根据sheet name
     * @param sheetName
     * @return
     */
    public ExcelSheet getSheet(String sheetName) {
        return new ExcelSheet(this.hssfWorkbook.getSheet(sheetName));
    }

    /**
     * 获取工作簿 根据
     * @param sheetIndex
     * @return
     */
    public ExcelSheet getSheet(int sheetIndex) {
        return new ExcelSheet(this.hssfWorkbook.getSheetAt(sheetIndex));
    }
    /**
     * 获取工作簿 根据
     * @return
     */
    public ExcelSheet getSheet() {
        return new ExcelSheet(this.hssfWorkbook.getSheetAt(0));
    }


    public HSSFWorkbook getHssfWorkbook() {
        return hssfWorkbook;
    }

    /**
     * 根据路径到处excel
     *
     * @param file
     * @throws IOException
     */
    public void saveExcel(String file) throws IOException {
        OutputStream out = new FileOutputStream(new File(file));
        this.hssfWorkbook.write(out);
    }

    /**
     * excel导出
     *
     * @param outputStream
     * @throws IOException
     */
    public void saveExcel(OutputStream outputStream) throws IOException {
        this.hssfWorkbook.write(outputStream);
    }
}
