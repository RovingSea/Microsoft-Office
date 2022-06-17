package com.rovingsea.study.microsoftOffice.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.listener.PageReadListener;
import com.alibaba.fastjson.JSON;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author Haixin Wu
 * @since 1.0
 */
public class EasyReadWriteTest {

    public static String PATH = "./src/main/resources/";

    private List<DemoData> dataList() {
        List<DemoData> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }

    @Test
    public void simpleWrite() {
        String fileName = PATH + "easyExcelTest.xlsx";
        EasyExcel.write(fileName, DemoData.class).sheet("1").doWrite(dataList());
    }

    @Test
    public void simpleRead() {
        String file = PATH + "easyExcelTest.xlsx";
        EasyExcel.read(file, DemoData.class, new PageReadListener<DemoData>(dataList -> {
            for (DemoData data : dataList) {
                System.out.println(JSON.toJSONString(data));
            }
        })).sheet().doRead();
    }

}

