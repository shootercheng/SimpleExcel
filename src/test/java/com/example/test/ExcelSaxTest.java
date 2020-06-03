package com.example.test;

import com.example.ExcelUtil;
import com.example.test.handlers.DefineRowHandler;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * @author James
 */
public class ExcelSaxTest {

    @Test
    public void testExcel2007() {
        String filePath = "file/test2007.xlsx";
        try (InputStream inputStream = new FileInputStream(filePath)) {
            DefineRowHandler defineRowHandler = new DefineRowHandler();
            ExcelUtil.readBySax(inputStream, 0, defineRowHandler);
            List<List<String>> dataList = defineRowHandler.getDataList();
            System.out.println(dataList);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testExcel2003() {
        String filePath = "file/test2003.xls";
        try (InputStream inputStream = new FileInputStream(filePath)) {
            DefineRowHandler defineRowHandler = new DefineRowHandler();
            ExcelUtil.readBySax(inputStream, 0, defineRowHandler);
            List<List<String>> dataList = defineRowHandler.getDataList();
            System.out.println(dataList);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
