package com.put.put;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSSF - 提供读写Microsoft Excel格式档案的功能。
 *
 * XSSF - 提供读写Microsoft Excel OOXML格式档案的功能。
 *
 * @author zhujiqian
 * @date 2020/10/27 20:33
 */
public class ExportExcelBase {
    private HSSFWorkbook hwb=null;
    private HSSFSheet sheet=null;

    public ExportExcelBase(HSSFWorkbook hwb,HSSFSheet sheet){
        this.hwb=hwb;
        this.sheet=sheet;
    }

    public HSSFWorkbook getHwb() {
        return hwb;
    }

    public void setHwb(HSSFWorkbook hwb) {
        this.hwb = hwb;
    }

    public HSSFSheet getSheet() {
        return sheet;
    }

    public void setSheet(HSSFSheet sheet) {
        this.sheet = sheet;
    }

    /**
     * 创建设置表格头
     */
    public void createNormalHead(String headString,int colSum){
        //创建表格标题行，第一行
        HSSFRow row=this.sheet.createRow(0);
        //创建指定行的列，第一列
        HSSFCell cell=row.createCell(0);
        //设置标题行默认行高
        row.setHeight((short) 500);
        //设置表格内容类型：0-数值型；1-字符串；2-公式型；3-空值；4-布尔型；5-错误
        cell.setCellType(1);
        //设置表格标题内容
        cell.setCellValue(new HSSFRichTextString(headString));
        // 指定合并区域
        this.sheet.addMergedRegion(new Region(0, (short)0, 0, (short)colSum));
        // 定义单元格格式，添加单元格表样式，并添加到工作簿
        HSSFCellStyle cellStyle=this.hwb.createCellStyle();
        // 设置单元格水平对齐居中类型
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 指定单元格垂直居中对齐
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 指定单元格自动换行
        cellStyle.setWrapText(true);
        //设置字体
        HSSFFont font=this.hwb.createFont();
        font.setBoldweight((short) 700);
        font.setFontName("宋体");
        font.setFontHeight((short) 300);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }

    /**
     *表单第二行
     * @param params
     * @param colSum
     */
    public void createNormalTwoRow(String[] params,int colSum){
        HSSFRow row1=this.sheet.createRow(1);
        row1.setHeight((short) 300);
        HSSFCell cell2=row1.createCell(0);
        cell2.setCellType(1);
        cell2.setCellValue(new HSSFRichTextString("统计时间"+params[0]+"至"+params[1]));
        this.sheet.addMergedRegion(new Region(1, (short) 0,1,(short)colSum));
        HSSFCellStyle cellStyle=this.hwb.createCellStyle();
        cellStyle.setAlignment((short) 2);
        cellStyle.setVerticalAlignment((short) 1);
        cellStyle.setWrapText(true);
        HSSFFont font=this.hwb.createFont();
        font.setBoldweight((short) 700);
        font.setFontName("宋体");
        font.setFontHeight((short) 250);
        cellStyle.setFont(font);
        cell2.setCellStyle(cellStyle);
    }


    /**
     * 表单内容
     * @param wb
     * @param row
     * @param col
     * @param align
     * @param val
     */
    public void cteateCell(HSSFWorkbook wb, HSSFRow row, int col, short align, String val) {
        HSSFCell cell = row.createCell(col);
        cell.setCellType(1);
        cell.setCellValue(new HSSFRichTextString(val));
        HSSFCellStyle cellstyle = wb.createCellStyle();
        cellstyle.setAlignment(align);
        cell.setCellStyle(cellstyle);
    }


    /**
     * 文档输出流
     * @param fileName
     */
    public void outputExcle(String fileName) {
        FileOutputStream fos = null;

        try {
            fos = new FileOutputStream(new File(fileName));
            this.hwb.write(fos);
            fos.close();
        } catch (FileNotFoundException var4) {
            var4.printStackTrace();
        } catch (IOException var5) {
            var5.printStackTrace();
        }

    }


}
