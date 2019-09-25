package com.alibaba.excel.event;

import com.alibaba.excel.context.AnalysisContext;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;


/**
 * @Auther: xyd
 * @Date: 2019/7/12 15:15
 * @Description: 复写ali 提供的 监听器 同时使用适配器模式提供默认的监听器
 */
@Data
public abstract class AnalysisEventListener<T> {

    private List<T> data = new ArrayList<>();

    public AnalysisEventListener() {
    }

    public abstract void invoke(T var1, AnalysisContext var2);

    public abstract void doAfterAllAnalysed(AnalysisContext var1);
}

