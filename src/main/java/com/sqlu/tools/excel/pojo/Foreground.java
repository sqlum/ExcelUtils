package com.sqlu.tools.excel.pojo;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.awt.*;

/**
 * 前景
 * @author: stonelu
 * @create: 2019-12-20 10:45
 **/
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class Foreground {
    private Color color;
}
