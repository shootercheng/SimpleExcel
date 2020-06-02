package com.example.test;

import com.example.ExcelUtil;
import com.example.test.handlers.DefineRowHandler;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

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
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
