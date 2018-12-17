package com.diaosichengxuyuan.tool.excel;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.IOUtils;

import java.io.*;
import java.util.List;

/**
 * @author liuhaipeng
 * @date 2018/12/17
 */
public class ReadUtil {

    /**
     * 读取excel中数据写入文件中，适用于小于1000行数据的场景
     *
     * @param readExcelFile 从该excel文件中读取数据，绝对路径
     * @param sheetNo       第几页，第一页是1
     * @param startLine     起始行，第一行是0
     * @param startColumn   起始列，第一列是0
     * @param endColumn     结束列，包含endColumn这一列
     * @param format        String.format格式的字符串
     * @param writeFile     写入该文件中，绝对路径
     * @param isClearFile   是否写入前先清空writeFile内容
     */
    public static void readExcelToFile(String readExcelFile, int sheetNo, int startLine, int startColumn, int endColumn,
                                       String format, String writeFile, boolean isClearFile) {
        File readFile0 = new File(readExcelFile);
        if(!readFile0.exists()) {
            System.out.println(String.format("需要读取的文件不存在，文件路径：%s", readFile0));
            return;
        }
        if(readFile0.isDirectory()) {
            System.out.println(String.format("需要读取的文件是一个目录，文件路径：%s", readFile0));
            return;
        }

        File writeFile0;
        try {
            writeFile0 = new File(writeFile);
            if(!writeFile0.exists()) {
                System.out.println(String.format("需要写入的文件不存在，创建文件，文件路径：%s", readFile0));
                writeFile0.createNewFile();
            } else if(writeFile0.isDirectory()) {
                System.out.println(String.format("需要写入的文件是一个目录，文件路径：%s", readFile0));
                return;
            } else if(isClearFile) {
                writeFile0.delete();
                writeFile0.createNewFile();
            }
        } catch(IOException e) {
            System.out.println(String.format("创建文件异常，文件路径：%s", readFile0));
            return;
        }

        FileInputStream inputStream = null;
        List<Object> oneSheetData;
        try {
            inputStream = new FileInputStream(readFile0);
            oneSheetData = EasyExcelFactory.read(inputStream, new Sheet(sheetNo, startLine));
        } catch(FileNotFoundException e) {
            e.printStackTrace();
            return;
        } finally {
            IOUtils.closeQuietly(inputStream);
        }

        if(CollectionUtils.isEmpty(oneSheetData)) {
            System.out.println(String.format("没有数据可读，文件路径：%s", readFile0));
            return;
        }

        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(writeFile0);
            for(int i = 0; i < oneSheetData.size(); i++) {
                String cellData = String.format(format, ((List)oneSheetData.get(i)).subList(startColumn, endColumn + 1)
                    .toArray());
                System.out.println(String.format("写入第%s行数据 ---> %s", i + 1, cellData));
                outputStream.write(cellData.getBytes("UTF-8"));
                if(i != oneSheetData.size() - 1) {
                    outputStream.write("\n".getBytes("UTF-8"));
                }
            }
        } catch(UnsupportedEncodingException e) {
            e.printStackTrace();
            return;
        } catch(IOException e) {
            e.printStackTrace();
            return;
        } finally {
            IOUtils.closeQuietly(outputStream);
        }
    }

    /**
     * 读取excel中数据写入文件中，适用于小于1000行数据的场景
     *
     * @param readExcelFile 从该excel文件中读取数据，绝对路径
     * @param sheetNo       第几页，第一页是1
     * @param startLine     起始行，第一行是0
     * @param startColumn   起始列，第一列是0
     * @param endColumn     结束列，包含endColumn这一列
     * @param format        String.format格式的字符串
     * @param writeFile     写入该文件中，绝对路径
     * @param isClearFile   是否写入前先清空writeFile内容
     */
    public static void saxReadExcelToFile(String readExcelFile, int sheetNo, int startLine, int startColumn,
                                          int endColumn, String format, String writeFile, boolean isClearFile) {

        /**
         * 方法内部类
         */
        class SaxModelListener extends AnalysisEventListener {
            private int i = 0;
            private OutputStream outputStream;
            private int startColumn;
            private int endColumn;
            private String format;

            private SaxModelListener(OutputStream outputStream, int startColumn, int endColumn, String format) {
                this.outputStream = outputStream;
                this.startColumn = startColumn;
                this.endColumn = endColumn;
                this.format = format;
            }

            @Override
            public void invoke(Object o, AnalysisContext analysisContext) {
                String cellData = String.format(format, ((List)o).subList(startColumn, endColumn + 1).toArray());
                try {
                    if(i != 0) {
                        outputStream.write("\n".getBytes("UTF-8"));
                    }
                    outputStream.write(cellData.getBytes("UTF-8"));
                } catch(Exception e) {
                    e.printStackTrace();
                    IOUtils.closeQuietly(outputStream);
                }
                System.out.println(String.format("写入第%s行数据 ---> %s", ++i, cellData));
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                System.out.println("写入文件内容结束");
            }
        }

        File readFile0 = new File(readExcelFile);
        if(!readFile0.exists()) {
            System.out.println(String.format("需要读取的文件不存在，文件路径：%s", readFile0));
            return;
        }
        if(readFile0.isDirectory()) {
            System.out.println(String.format("需要读取的文件是一个目录，文件路径：%s", readFile0));
            return;
        }

        File writeFile0;
        try {
            writeFile0 = new File(writeFile);
            if(!writeFile0.exists()) {
                System.out.println(String.format("需要写入的文件不存在，创建文件，文件路径：%s", readFile0));
                writeFile0.createNewFile();
            } else if(writeFile0.isDirectory()) {
                System.out.println(String.format("需要写入的文件是一个目录，文件路径：%s", readFile0));
                return;
            } else if(isClearFile) {
                writeFile0.delete();
                writeFile0.createNewFile();
            }
        } catch(IOException e) {
            System.out.println(String.format("创建文件异常，文件路径：%s", readFile0));
            return;
        }

        FileInputStream inputStream = null;
        FileOutputStream outputStream = null;
        try {
            inputStream = new FileInputStream(readFile0);
            outputStream = new FileOutputStream(writeFile0);
            EasyExcelFactory.readBySax(inputStream, new Sheet(sheetNo, startLine),
                new SaxModelListener(outputStream, startColumn, endColumn, format));
        } catch(FileNotFoundException e) {
            e.printStackTrace();
            return;
        } finally {
            IOUtils.closeQuietly(inputStream);
            IOUtils.closeQuietly(outputStream);
        }
    }

}
