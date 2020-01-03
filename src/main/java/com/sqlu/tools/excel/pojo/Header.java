package com.sqlu.tools.excel.pojo;

import lombok.Data;

import java.util.*;

/**
 * 表头
 * @author: stonelu
 * @create: 2019-12-20 11:03
 **/
@Data
public class Header {
    private static final Header INSTANCE = new Header();
    private Header() {}
    private Header(int rowCount) {
        this.headerRows = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            this.headerRows.add(new HeaderRowBuilder());
        }

        colIdxWidthMap = new HashMap<>();
    }

    List<HeaderRowBuilder> headerRows;
    /** 列索引-宽度 map */
    private Map<Integer, Integer> colIdxWidthMap;
    private Font font;
    private Border border;
    private Foreground foreground;
    private Align align;

    public static HeaderBuilder builder(int rowCount) {
        return INSTANCE.new HeaderBuilder(rowCount);
    }

    public class HeaderBuilder {
        Header header;

        private HeaderBuilder(int rowCount) {
            header = new Header(rowCount);
        }

        public HeaderBuilder font(Font font) {
            header.font = font;
            return this;
        }

        public HeaderBuilder border(Border border) {
            header.border = border;
            return this;
        }

        public HeaderBuilder foreground(Foreground foreground) {
            header.foreground = foreground;
            return this;
        }

        public HeaderBuilder align(Align align) {
            header.align = align;
            return this;
        }

        /**
         * 获取指定行
         * @param rowIdx 行号，从0开始
         * @return
         */
        public HeaderRowBuilder row(int rowIdx) {
            if (rowIdx < 0 || rowIdx >= header.getHeaderRows().size()) {
                throw new IllegalArgumentException("行号超出范围：" + rowIdx);
            }

            return header.getHeaderRows().get(rowIdx);
        }

        /**
         * 添加列宽度
         * @param colIdx
         * @param colWidth
         * @return
         */
        public HeaderBuilder colWidth(int colIdx, int colWidth) {
            header.colIdxWidthMap.put(colIdx, colWidth);
            return this;
        }

        public Header build() {
            return header;
        }
    }

    public class HeaderRowBuilder {
        /** 根据 里氏替换原则，此处只需要存储父类Cell即可 */
        private List<Cell> cells;
        public HeaderRowBuilder() {
            cells = new LinkedList<>();
        }

        public List<Cell> build() {
            return this.cells;
        }

        public HeaderRowBuilder append(Cell cell) {
            this.cells.add(cell);
            return this;
        }

        public HeaderRowBuilder append(String content) {
            Cell cell = new Cell();
            cell.setContent(content);
            this.cells.add(cell);
            return this;
        }

        public HeaderRowBuilder appendWithColWidth(String content, int colWidth) {
            Cell cell = new Cell();
            cell.setContent(content);
            colIdxWidthMap.put(cells.size(), colWidth);
            this.cells.add(cell);
            return this;
        }

        public HeaderRowBuilder append(String content, int skipCellCount) {
            Cell cell = new Cell();
            cell.setContent(content);
            cell.setSkipCellCount(skipCellCount);
            this.cells.add(cell);
            return this;
        }

        public HeaderRowBuilder append(String content, Integer mergeRowCount, Integer mergeColCount) {
            MergeCell mergeCell = new MergeCell();
            mergeCell.setContent(content);
            mergeCell.setMergeRowCount(mergeRowCount);
            mergeCell.setMergeColCount(mergeColCount);

            this.cells.add(mergeCell);
            return this;
        }

        public HeaderRowBuilder append(String content, Integer colWidth, Integer mergeRowCount, Integer mergeColCount) {
            MergeCell mergeCell = new MergeCell();
            mergeCell.setContent(content);
            mergeCell.setMergeRowCount(mergeRowCount);
            mergeCell.setMergeColCount(mergeColCount);

            if (notNull(colWidth)) {
                colIdxWidthMap.put(cells.size(), colWidth);
            }

            this.cells.add(mergeCell);
            return this;
        }
    }

    public boolean isNull(Object o) {return null == o;}
    public boolean notNull(Object o) {return !isNull(o);}
}

