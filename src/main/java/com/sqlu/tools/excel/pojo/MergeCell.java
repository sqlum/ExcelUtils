package com.sqlu.tools.excel.pojo;

import lombok.Data;

/**
 * 单元格：支持合并
 * @author: stonelu
 * @create: 2019-12-20 11:13
 **/
@Data
public class MergeCell extends Cell{
    /** 合并行数 */
    private Integer mergeRowCount;
    /** 合并列数 */
    private Integer mergeColCount;
}
