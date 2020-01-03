package com.sqlu.tools.excel.pojo;

import lombok.Data;

/**
 * 字体
 * @author: stonelu
 * @create: 2019-12-20 10:31
 **/
@Data
public class Font {
    /** 大小 */
    private Double size;
    /** 名称 */
    private String fontName;
    /** 是否加粗 */
    private Boolean isBold;
}
