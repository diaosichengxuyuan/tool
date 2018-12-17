package com.diaosichengxuyuan.tool.excel;

import org.junit.Test;

/**
 * @author liuhaipeng
 * @date 2018/12/17
 */
public class ReadUtilTest {

    @Test
    public void testReadExcelToFile() {
        String from = "/Users/liuhaipeng/Documents/code/IdeaProject/tool/com.diaosichengxuyuan.tool.parent/com"
            + ".diaosichengxuyuan.tool.excel/src/test/resources/from.xlsx";
        String formatStr = "姓名：%s，年龄：%s，身高：%s";
        String to = "/Users/liuhaipeng/Downloads/to.txt";

        long startTime = System.currentTimeMillis();
        ReadUtil.readExcelToFile(from, 2, 2, 2, 4, formatStr, to, true);
        System.out.println(String.format("readExcelToFile耗时%s毫秒", System.currentTimeMillis() - startTime));
    }

    @Test
    public void testSaxReadExcelToFile() {
        String from = "/Users/liuhaipeng/Documents/code/IdeaProject/tool/com.diaosichengxuyuan.tool.parent/com"
            + ".diaosichengxuyuan.tool.excel/src/test/resources/from.xlsx";
        String formatStr = "姓名：%s，年龄：%s，身高：%s";
        String to = "/Users/liuhaipeng/Downloads/to.txt";

        long startTime = System.currentTimeMillis();
        ReadUtil.saxReadExcelToFile(from, 2, 2, 2, 4, formatStr, to, true);
        System.out.println(String.format("readExcelToFile耗时%s毫秒", System.currentTimeMillis() - startTime));
    }

}
