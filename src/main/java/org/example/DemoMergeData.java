package org.example;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;



@Getter
@Setter
@EqualsAndHashCode
class DemoMergeData {
    @ExcelProperty("id")
    int id;
    @ExcelProperty("姓名")
    String sub;
    @ExcelProperty("分数")
    String date;
    DemoMergeData(int id, String sub,String date) {
        this.id =id;
        this.sub =sub;
        this.date =date;
    }
    DemoMergeData() {

    }
    public static   DemoMergeData builder() {
        return new DemoMergeData();
    }
    DemoMergeData id(int id) {
        this.id = id;
        return  this;
    }
    DemoMergeData sub(String sub) {
        this.sub =sub;
        return this;
    }
    DemoMergeData date(String date) {
        this.date = date;
        return this;
    }
    DemoMergeData build() {
        return this;
    }
}
