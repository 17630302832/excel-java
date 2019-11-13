package com.worldline.autotest.e2e.data.excel;

import com.worldline.autotest.e2e.data.IFileDataSaver;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class CFileDataSaver implements IFileDataSaver {
    static final Logger logger = LoggerFactory.getLogger(CFileDataSaver.class);

    @Override
    public void saveData(File resource, List<String> headers, List<Map<String, Object>> records) throws IOException {
        //创建excel工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建工作表sheet
        XSSFSheet sheet = workbook.createSheet();
        //创建第一行
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = null;
        FileOutputStream stream = null;

        //插入第一行数据的表头
        for (int i = 0; i < headers.size(); i++) {
            cell = row.createCell(i);
            cell.setCellValue(headers.get(i));
        }
        //插入数据
        for (int i = 1; i < records.size(); i++) {
            XSSFRow nrow = sheet.createRow(i);
            for (int j = 0; j < headers.size(); j++) {
                XSSFCell ncell = nrow.createCell(j);
                if (!records.get(i).isEmpty()) {
                    if (!headers.get(j).isEmpty()) {
                        ncell.setCellValue(records.get(i).get(headers.get(j)).toString());
                    }
                }
            }
        }
        try {
            stream = FileUtils.openOutputStream(resource);
            workbook.write(stream);
            logger.info("导入成功");
        } finally {
            stream.close();
        }
    }
}
