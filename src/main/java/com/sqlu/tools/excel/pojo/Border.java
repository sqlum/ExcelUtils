package com.sqlu.tools.excel.pojo;

import lombok.Data;

/**
 * 边框
 * @author: stonelu
 * @create: 2019-12-20 10:40
 **/
@Data
public class Border {
    private Boolean containTop;
    private Boolean containBottom;
    private Boolean containLeft;
    private Boolean containRight;

    public Border(boolean allBorder) {
        if (allBorder) {
            this.containTop = true;
            this.containBottom = true;
            this.containLeft = true;
            this.containRight = true;
        }
    }
}
