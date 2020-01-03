package com.sqlu.tools.excel;

import com.sqlu.tools.excel.poi.PoiUtil;
import com.sqlu.tools.excel.pojo.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.awt.*;

/**
 * @author: stonelu
 * @create: 2020-01-03 15:00
 **/
public class PoiUtilTest {

    @Test
    public void test() {
        PoiUtil poiUtil = PoiUtil.INSTANCE;

        // 创建workbook
        XSSFWorkbook workbook = poiUtil.createXssfWorkbook();
        // 创建sheet，可以给sheet标签指定颜色
        XSSFSheet gradeDataSheet = poiUtil.createSheet(workbook, "年级得分率", new Color(81, 204, 186));
        XSSFSheet replenishSheet = poiUtil.createSheet(workbook, "补充", new Color(255, 112, 10));

        int headerLineCount = 2;

        // 表头：
        Header.HeaderBuilder headerBuilder = Header.builder(headerLineCount);

        // 样式
        headerBuilder.foreground(Foreground.builder().color(new Color(50, 125, 180)).build());
        headerBuilder.border(new Border(true));
        headerBuilder.align(Align.builder().isVerticalCenter(true).isHorizontalCenter(true).build());

        // 第一行
        headerBuilder.row(0).append("编号", 5, 2, 1)
                .append("年级得分率", 15, 2, 1)
                .append("年级", 1, 2)
                .append("一年级 1班", 1, 2)
                .append("一年级 2班", 1, 2);
        // 第二行
        headerBuilder.row(1)
                // 可在循环中按需append
                .append("人数", 2).append("占比")
                .append("人数").append("占比")
                .append("人数").append("占比");

        // 渲染表头
        poiUtil.drawHeader(workbook, gradeDataSheet, headerBuilder.build());




        // 表格体：
        Content.ContentBuilder contentBuilder = Content.builder();

        // 样式
        contentBuilder.align(Align.builder().isHorizontalCenter(true).isVerticalCenter(true).build());
        contentBuilder.border(new Border(true));

        contentBuilder.newRow()
                // 编号、年级得分率 分别合并2行1列
                .append(1, 2, 1)
                .append("87.5%", 2, 1)
                // 表格体中的第一行内容
                .append(7)
                .append("87.5%")
                .append(2)
                .append("100%")
                .append(5)
                .append("83.33%");

        contentBuilder.newRow()
                // 表格体的第二行前两列被占用了，故需跳过
                .append(1, 2)
                .append("12.5%")
                .append(0)
                .append("0%")
                .append(1)
                .append("16.7%");

        contentBuilder.newRow()
                // 编号、年级得分率 分别合并2行1列
                .append(2, 2, 1)
                .append("75%", 2, 1)
                // 表格体中的第一行内容
                .append(6)
                .append("75%")
                .append(2)
                .append("100%")
                .append(4)
                .append("66.7%");

        contentBuilder.newRow()
                // 表格体的第二行前两列被占用了，故需跳过
                .append(2, 2)
                .append("25%")
                .append(0)
                .append("0%")
                .append(0)
                .append("0%");
        // 绘制表格体
        poiUtil.drawContent(workbook, gradeDataSheet, headerLineCount, contentBuilder.build());

        // 输出到文件
        poiUtil.writeExcel2File("/Users/Stonelu/Desktop/out.xlsx", workbook);
    }


}
