package com.cty.easyexceldemo;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * @author ：Mr.chen
 * @date ：Created in 2022/4/22 22:30
 * @description：合并单元格
 * @modified By：
 * @version: $
 */
public class ExcelFillCellMergeStrategy implements CellWriteHandler {
    // 需要从第几行开始合并，0表示第1行
    private int mergeRowIndex = 1;
    // 合并的哪些列，比如为4时，当前行id和上一行id相同则合并前五列
    private int mergeColumnRegion = 4;

    public ExcelFillCellMergeStrategy() {
    }


    public ExcelFillCellMergeStrategy(int mergeColumnRegion) {
        this.mergeColumnRegion = mergeColumnRegion;
    }

    public ExcelFillCellMergeStrategy(int mergeRowIndex, int mergeColumnRegion) {
        this.mergeRowIndex = mergeRowIndex;
        this.mergeColumnRegion = mergeColumnRegion;
    }
    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer columnIndex, Integer relativeRowIndex, Boolean isHead) {

    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // 隐藏id列
//        writeSheetHolder.getSheet().setColumnHidden(0, true);
    }

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, Cell cell, Head head, Integer integer, Boolean aBoolean) {

    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> list, Cell cell, Head head, Integer integer, Boolean aBoolean) {
        //当前行
        int curRowIndex = cell.getRowIndex();
        //当前列
        int curColIndex = cell.getColumnIndex();

        /**
         * 样式在这里设置
         */
        Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellStyle(cellStyle);

        if (curRowIndex > mergeRowIndex) {
            for (int i = 0; i < mergeColumnRegion; i++) {
                if (curColIndex <= mergeColumnRegion) {
                    mergeWithPreviousRow(writeSheetHolder, cell, curRowIndex, curColIndex);
                    break;
                }
            }
        }
    }

    /**
     * 当前单元格向上合并：当前行的id和上一行的id相同则合并前面（mergeColumnRegion+1）列
     *
     * @param writeSheetHolder
     * @param cell             当前单元格
     * @param curRowIndex      当前行
     * @param curColIndex      当前列
     */
    private void mergeWithPreviousRow(WriteSheetHolder writeSheetHolder, Cell cell, int curRowIndex, int curColIndex) {
        // 当前行的第一个Cell
        Cell curFirstCell = cell.getSheet().getRow(curRowIndex).getCell(0);
        Object curFirstData = curFirstCell.getCellTypeEnum() == CellType.STRING ? curFirstCell.getStringCellValue() : curFirstCell.getNumericCellValue();
        // 上一行的第一个Cell
        Cell preFirstCell = cell.getSheet().getRow(curRowIndex - 1).getCell(0);
        Object preFirstData = preFirstCell.getCellTypeEnum() == CellType.STRING ? preFirstCell.getStringCellValue() : preFirstCell.getNumericCellValue();

        // 当前行的id和上一行的id相同则合并前面（mergeColumnRegion+1）列
        if (curFirstData.equals(preFirstData)) {
            Sheet sheet = writeSheetHolder.getSheet();
            List<CellRangeAddress> mergeRegions = sheet.getMergedRegions();
            boolean isMerged = false;
            for (int i = 0; i < mergeRegions.size() && !isMerged; i++) {
                CellRangeAddress cellRangeAddr = mergeRegions.get(i);
                // 若上一个单元格已经被合并，则先移出原有的合并单元，再重新添加合并单元
                if (cellRangeAddr.isInRange(curRowIndex - 1, curColIndex)) {
                    sheet.removeMergedRegion(i);
                    cellRangeAddr.setLastRow(curRowIndex);
                    sheet.addMergedRegion(cellRangeAddr);
                    isMerged = true;
                }
            }
            // 若上一个单元格未被合并，则新增合并单元
            if (!isMerged) {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(curRowIndex - 1, curRowIndex, curColIndex, curColIndex);
                sheet.addMergedRegion(cellRangeAddress);
            }
        }
    }

}
