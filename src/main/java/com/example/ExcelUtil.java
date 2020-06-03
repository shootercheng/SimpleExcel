package com.example;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ExcelUtil {

    /**
     * 调用解析
     * @param inputStream excel文件输入流
     * @param sheetIndex sheet表0,1,2
     * @param rowDataHandler 解析数据回调接口
     */
    public static void readBySax(InputStream inputStream, int sheetIndex, RowDataHandler rowDataHandler) {

        BufferedInputStream bif = new BufferedInputStream(inputStream);

        if (isXlsx(bif)) {
            try {
                read07BySax(bif, sheetIndex,rowDataHandler);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (OpenXML4JException e) {
                e.printStackTrace();
            }
        } else {
            read03BySax(bif, sheetIndex,rowDataHandler);
        }
    }
    private static boolean isXlsx(InputStream in) {
        try {
            return FileMagic.valueOf(in) == FileMagic.OOXML;
        } catch (IOException e) {
            e.printStackTrace();
        }

        return false;
    }


    private static Excel03SaxReader read03BySax(InputStream in, Integer sheetIndex,RowDataHandler rowDataHandler) {

        return new Excel03SaxReader(rowDataHandler).read(in, sheetIndex);
    }

    private static Excel07SaxReader read07BySax(InputStream in, Integer sheetIndex,RowDataHandler rowDataHandler) throws IOException, OpenXML4JException {

        return new Excel07SaxReader(rowDataHandler).read(in, sheetIndex);
    }
}
