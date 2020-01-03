package com.sqlu.tools.excel.pojo;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 对齐
 * @author: stonelu
 * @create: 2019-12-20 10:47
 **/
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class Align {
    /** 水平对齐 */
    private Boolean isLeft;
    private Boolean isHorizontalCenter;
    private Boolean isRight;

    /** 垂直对齐 */
    private Boolean isTop;
    private Boolean isVerticalCenter;
    private Boolean isBottom;
}
