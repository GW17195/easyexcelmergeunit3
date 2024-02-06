package org.example;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class MyTest {
    @Test
    public void testWrite2() throws IOException {
        // 数据就不初始化了
        List<DemoMergeData> resultList = new ArrayList<>();
        resultList.add(DemoMergeData.builder().id(1).sub("张胜男").date("12").build());
        resultList.add(DemoMergeData.builder().id(1).sub("李四").date("224").build());
        resultList.add(DemoMergeData.builder().id(3).sub("王五").date("224").build());
        resultList.add(DemoMergeData.builder().id(4).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(5).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(5).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(8).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(8).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(9).sub("陈琪").date("224").build());
        resultList.add(DemoMergeData.builder().id(10).sub("小白").date("241").build());
        resultList.add(DemoMergeData.builder().id(11).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(12).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(12).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(12).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(13).sub("小黑").date("241").build());
        // 设置文件名称
        String fileName = "D:\\Users\\dell\\java_project\\easyexcelmergeunit3\\src\\main\\resources\\t3.xlsx";
        File file = new File(fileName);
        if (!file.exists()) {
            file.createNewFile();
        }

        //  sheet名称
        EasyExcel.write(fileName, DemoMergeData.class)
                .autoCloseStream(Boolean.TRUE)
                .registerWriteHandler(new MultiColumnMergeStrategy(resultList.size(),0,1))
                .sheet("测试导出").doWrite(resultList);
    }

    @Test
    public void testWrite1() throws IOException {
        int[] mergeColumnIndex = {0,1};
        // 需要从第几行开始合并
        int mergeRowIndex = 1;
        // 数据就不初始化了
        List<DemoMergeData> resultList = new ArrayList<>();
        resultList.add(DemoMergeData.builder().id(1).sub("张胜男").date("12").build());
        resultList.add(DemoMergeData.builder().id(1).sub("李四").date("224").build());
        resultList.add(DemoMergeData.builder().id(3).sub("王五").date("224").build());
        resultList.add(DemoMergeData.builder().id(4).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(5).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(5).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(8).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(8).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(9).sub("陈琪").date("224").build());
        resultList.add(DemoMergeData.builder().id(10).sub("小白").date("241").build());
        resultList.add(DemoMergeData.builder().id(11).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(12).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(12).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(12).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(13).sub("小黑").date("241").build());

        // 设置文件名称
        String fileName = "D:\\Users\\dell\\java_project\\easyexcelmergeunit3\\src\\main\\resources\\t2.xlsx";
        File file = new File(fileName);
        if (!file.exists()) {
            file.createNewFile();
        }

        //  sheet名称
        EasyExcel.write(fileName, DemoMergeData.class)
                .autoCloseStream(Boolean.TRUE)
                .registerWriteHandler(new ExcelFillCellMergeStrategy(mergeRowIndex,mergeColumnIndex))
                .sheet("测试导出").doWrite(resultList);
    }

    @Test
    public void testWrite() throws IOException {
        List<DemoMergeData> resultList = new ArrayList<>();
        resultList.add(DemoMergeData.builder().id(1).sub("张胜男").date("12").build());
        resultList.add(DemoMergeData.builder().id(1).sub("李四").date("224").build());
        resultList.add(DemoMergeData.builder().id(3).sub("王五").date("224").build());
        resultList.add(DemoMergeData.builder().id(4).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(5).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(5).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(8).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(8).sub("赵柳").date("224").build());
        resultList.add(DemoMergeData.builder().id(9).sub("陈琪").date("224").build());
        resultList.add(DemoMergeData.builder().id(10).sub("小白").date("241").build());
        resultList.add(DemoMergeData.builder().id(11).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(12).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(12).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(12).sub("小黑").date("241").build());
        resultList.add(DemoMergeData.builder().id(13).sub("小黑").date("241").build());
        //  设置文件名称
        String fileName = "D:\\Users\\dell\\java_project\\easyexcelmergeunit3\\src\\main\\resources\\t1.xlsx";
        File file = new File(fileName);
        if (!file.exists()) {
            file.createNewFile();
        }

        //  sheet名称
        EasyExcel.write(fileName, DemoMergeData.class)
                .autoCloseStream(Boolean.TRUE)
                .sheet("测试导出").doWrite(resultList);
    }
}
