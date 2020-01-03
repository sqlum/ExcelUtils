package com.sqlu.tools.excel.pojo;

import lombok.Data;

import java.awt.*;
import java.util.List;
import java.util.*;

/**
 * @author: stonelu
 * @create: 2019-12-21 17:14
 **/
@Data
public class Content {
    private List<ContentRowBuilder> contentRows;
    /** 行索引-行高 mao */
    private Map<Integer, Integer> rowIdxHeightMap;
    private Font font;
    private Border border;
    private Foreground foreground;
    private Align align;

    private static final Content INSTANCE =  new Content();

    private Content() {
        this.contentRows = new LinkedList<>();
        this.rowIdxHeightMap = new HashMap<>();
    }

    public static ContentBuilder builder() {
        return INSTANCE.new ContentBuilder();
    }


    public class ContentBuilder {
        private Content content;

        public ContentBuilder() {
            this.content = new Content();
        }

        public ContentBuilder font(Font font) {
            this.content.setFont(font);
            return this;
        }

        public ContentBuilder border(Border border) {
            this.content.setBorder(border);
            return this;
        }

        public ContentBuilder foreground(Foreground foreground) {
            this.content.setForeground(foreground);
            return this;
        }

        public ContentBuilder align(Align align) {
            this.content.setAlign(align);
            return this;
        }

        public ContentRowBuilder newRow() {
            ContentRowBuilder contentRowBuilder = new ContentRowBuilder();
            content.getContentRows().add(contentRowBuilder);
            return contentRowBuilder;
        }

        /**
         * 新建行，并设置行高度
         * @param rowHeight
         * @return
         */
        public ContentRowBuilder newRowWithRowHeight(int rowHeight) {
            ContentRowBuilder builder = newRow();
            this.content.getRowIdxHeightMap().put(getCurrRowIdx(), rowHeight);
            return builder;
        }

        private int getCurrRowIdx() {
            return this.content.contentRows.size();
        }

        public Content build() {
            return this.content;
        }
    }


    public class ContentRowBuilder {
        /** 根据 里氏替换原则，此处只需要存储父类Cell即可 */
        private List<Cell> cells;

        public ContentRowBuilder() {
            cells = new LinkedList<>();
        }

        public List<Cell> build() {
            return this.cells;
        }

        public ContentRowBuilder append(Object content) {
            return this.append(content, 0);
        }

        public ContentRowBuilder appendWithColWidth(String content, int colWidth) {
            Cell cell = new Cell();
            cell.setContent(content);

            this.cells.add(cell);
            return this;
        }

        public ContentRowBuilder append(Object content, int skipCellCount) {
            return this.append(content,skipCellCount, null);
        }

        public ContentRowBuilder append(Object content, int skipCellCount, Align align) {
            Cell cell = new Cell();
            cell.setContent(content);
            cell.setSkipCellCount(skipCellCount);
            cell.setAlign(align);

            this.cells.add(cell);
            return this;
        }

        /**
         * 添加内容
         * @param content
         * @param foregroundColor 前景色
         * @return
         */
        public ContentRowBuilder append(Object content, Color foregroundColor) {
            Cell cell = new Cell();
            cell.setContent(content);

            // 设置前景色
            if (Objects.nonNull(foregroundColor)) {
                cell.setForeground(Foreground.builder().color(foregroundColor).build());
            }

            this.cells.add(cell);
            return this;
        }

        public ContentRowBuilder append(Object content, Align align) {
            Cell cell = new Cell();
            cell.setContent(content);
            cell.setAlign(align);

            this.cells.add(cell);
            return this;
        }

        public ContentRowBuilder append(Object content, Integer mergeRowCount, Integer mergeColCount) {
            return append(content, false, null, mergeRowCount, mergeColCount);
        }

        public ContentRowBuilder append(Object content, Integer mergeRowCount, Integer mergeColCount, Integer skipCellCount) {
            MergeCell mergeCell = new MergeCell();
            mergeCell.setContent(content);
            mergeCell.setMergeRowCount(mergeRowCount);
            mergeCell.setMergeColCount(mergeColCount);

            // 是否自动换行
            mergeCell.setSkipCellCount(skipCellCount);

            this.cells.add(mergeCell);
            return this;
        }

        /**
         * 添加内容
         * @param content
         * @param isAutoWrapText 是否自动换行
         * @param align
         * @param mergeRowCount 合并行数
         * @param mergeColCount 合并列数
         * @return
         */
        public ContentRowBuilder append(Object content, boolean isAutoWrapText, Align align, Integer mergeRowCount, Integer mergeColCount) {
            return this.append(content, isAutoWrapText, align, mergeRowCount, mergeColCount, null);
        }

        public ContentRowBuilder append(Object content, boolean isAutoWrapText, Align align, Integer mergeRowCount, Integer mergeColCount, Integer skipCellCount) {
            MergeCell mergeCell = new MergeCell();
            mergeCell.setContent(content);
            mergeCell.setMergeRowCount(mergeRowCount);
            mergeCell.setMergeColCount(mergeColCount);

            mergeCell.setSkipCellCount(skipCellCount);

            // 是否自动换行
            mergeCell.setIsAutoWrapText(isAutoWrapText);
            // 对齐
            mergeCell.setAlign(align);

            this.cells.add(mergeCell);
            return this;
        }

        /**
         * 添加一行中的单元格内容
         * @param content 内容
         * @param isAutoWrapText 是否自动换行
         * @return
         */
        public ContentRowBuilder append(Object content, boolean isAutoWrapText) {
            return this.append(content, isAutoWrapText, null);
        }

        /**
         * 添加一行中的单元格内容
         * @param content 内容
         * @param isAutoWrapText 是否自动换行(注意：建议给行高设个值，否则此功能可能不生效)
         * @param foregroundColor 前景色
         * @return
         */
        public ContentRowBuilder append(Object content, boolean isAutoWrapText, Color foregroundColor) {
            Cell cell = new Cell();
            cell.setContent(content);
            cell.setIsAutoWrapText(isAutoWrapText);

            // 设置前景色
            if (Objects.nonNull(foregroundColor)) {
                cell.setForeground(Foreground.builder().color(foregroundColor).build());
            }

            this.cells.add(cell);
            return this;
        }
    }
}
