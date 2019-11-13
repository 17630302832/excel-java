package com.worldline.autotest.e2e.data.excel;

import com.worldline.autotest.e2e.data.IFileDataLoader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class CFileDataLoader implements IFileDataLoader {
    static final Logger logger = LoggerFactory.getLogger(CFileDataLoader.class);

    @Override
    public List<Map<String, Object>> loadData(File resource) throws IOException {
        Workbook wb = null;
        ArrayList<Map<String, Object>> result = null;
        try {
            wb = WorkbookFactory.create(resource);
            result = readExcel(wb, 0, 1, 0);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
        return result;
    }

    public ArrayList<Map<String, Object>> readExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        List<String> headers = readHeaders(sheet);
        Row row = null;
        ArrayList<Map<String, Object>> result = new ArrayList<Map<String, Object>>();
        for (int i = startReadLine; i < sheet.getLastRowNum() - tailLine + 1; i++) {
            row = sheet.getRow(i);
            if (row == null) continue;
            Map<String, Object> map = new HashMap<String, Object>();
            for (Cell c : row) {
                if (c.getColumnIndex() > headers.size() - 1) break;
                String returnStr = "";
                boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
                //判断是否具有合并单元格
                if (isMerge) {
                    String rs = getMergedRegionValue(sheet, row.getRowNum(), c.getColumnIndex());
                    returnStr = rs;
                } else {
                    returnStr = c.getRichStringCellValue().getString().trim();
                }
                map.put(String.valueOf(sheet.getRow(0).getCell(c.getColumnIndex())), returnStr);
            }
            result.add(map);
        }
        return result;

    }

    /**
     * 读取Excel表头
     *
     * @param sheet
     * @return
     */
    public List<String> readHeaders(Sheet sheet) {
        org.apache.poi.ss.usermodel.Row row = null;
        row = sheet.getRow(0);
        List<String> headers = new ArrayList<>();
        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (String.valueOf(row.getCell(i)).isEmpty()) {
                logger.info("表头信息不能为空,中断读取");
                break;
            } else {
                headers.add(String.valueOf(row.getCell(i)));
            }
        }
        return headers;
    }

    /**
     * 获取合并单元格的值
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell);
                }
            }
        }

        return null;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     *
     * @param sheet
     * @param row    行下标
     * @param column 列下标
     * @return
     */
    public boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    public String getCellValue(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        }
        return "";
    }
}

