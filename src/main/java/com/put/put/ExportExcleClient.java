package com.put.put;


import com.put.data.Datas;
import org.apache.poi.hssf.usermodel.*;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * TODO
 *
 * @author zhujiqian
 * @date 2020/10/27 20:24
 */
public class ExportExcleClient {
    private HSSFWorkbook hwb=null;
    private HSSFSheet sheet=null;

    ExportExcelBase exportExcel = null;
    SimpleDateFormat df = new SimpleDateFormat("yyyy年MM月dd日");

    public ExportExcleClient() {
        this.hwb = new HSSFWorkbook();
        this.exportExcel = new ExportExcelBase(this.hwb, this.sheet);
    }

    /**
     * 导出Excel
     * @return
     */
    public String exportExcel() {
        String a = this.df.format(new Date()) + "对账单集合.xls";
        String b = "D:\\批量处理对账单\\对账单集处理结果\\" + a;
        this.exportExcel.outputExcle(b);
        return b;
    }

    /**
     * 设置导出格式
     * @param data
     * @return
     */
    public String alldata(List<Datas> data) {
        if (data.size() != 0) {
            this.sheet = this.exportExcel.getHwb().createSheet("对账单集合");
            this.exportExcel.setSheet(this.sheet);
            int number = 2;

            for(int i = 0; i < number; ++i) {
                this.sheet.setColumnWidth(i, 8000);
            }

            HSSFCellStyle cellStyle = this.hwb.createCellStyle();
            cellStyle.setAlignment((short)2);
            cellStyle.setVerticalAlignment((short)1);
            cellStyle.setWrapText(true);
            HSSFFont font = this.hwb.createFont();
            font.setBoldweight((short)700);
            font.setFontName("宋体");
            font.setFontHeight((short)200);
            cellStyle.setFont(font);
            this.exportExcel.createNormalHead("对账单整合表", number);
            String[] params = new String[]{this.df.format(new Date()), this.df.format(new Date())};
            this.exportExcel.createNormalTwoRow(params, number);
            HSSFRow row2 = this.sheet.createRow(2);
            HSSFCell cell0 = row2.createCell(0);
            cell0.setCellStyle(cellStyle);
            cell0.setCellValue(new HSSFRichTextString("组合名"));
            HSSFCell cell1 = row2.createCell(1);
            cell1.setCellStyle(cellStyle);
            cell1.setCellValue(new HSSFRichTextString("保证金占用金额"));
            HSSFCell cell2 = row2.createCell(2);
            cell2.setCellStyle(cellStyle);
            cell2.setCellValue(new HSSFRichTextString("可用资金额"));

            for(int i = 0; i < data.size(); ++i) {
                System.out.println("==============" + ((Datas)data.get(i)).getName() + " " + ((Datas)data.get(i)).getMargin() + " " + ((Datas)data.get(i)).getAvaFunds());
                HSSFRow row = this.sheet.createRow((short)i + 3);
                this.exportExcel.cteateCell(this.hwb, row, 0, (short)6, ((Datas)data.get(i)).getName());
                this.exportExcel.cteateCell(this.hwb, row, 1, (short)6, ((Datas)data.get(i)).getMargin());
                this.exportExcel.cteateCell(this.hwb, row, 2, (short)6, ((Datas)data.get(i)).getAvaFunds());
            }
        }

        return "";
    }
}
