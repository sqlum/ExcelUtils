package com.sqlu.tools.excel.pojo;

import lombok.Data;

/**
 * 单元格
 * @author: stonelu
 * @create: 2019-12-20 10:52
 **/
@Data
public class Cell {
    /** 内容 */
    private Object content;
    /** 是否自动换行，支持\n强制换行 */
    private Boolean isAutoWrapText;

    /** 宽度 */
    private Double width;
    /** 高度 */
    private Double height;

    /** 跳过几个(水平)单元格 */
    private Integer skipCellCount;


    private Font font;
    private Border border;
    private Foreground foreground;
    private Align align;
}
