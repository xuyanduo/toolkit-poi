package com.alibaba.excel.event.adapter;

import annotation.FieldNotNull;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.util.StringUtils;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

import service.BeanToMap;

/**
 * @Auther: xyd
 * @Date: 2019/7/12 14:41
 * @Description:
 */
@Slf4j
@Data
public class DefaultAnalysisEventListenerAdapter<T> extends AnalysisEventListener<T> {

    private List<T> noValidData = new ArrayList<>();

    @Override
    public void invoke(T o, AnalysisContext analysisContext) {
        if (isEmpty(o)) {
            noValidData.add(o);
            log.warn("无效数据：{}", o);
        } else {
            this.getData().add(o);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }

    /**
     * 必填字段的空判断
     *
     * @param o pojo
     * @return boolean
     */
    private boolean isEmpty(T o) {
        Field[] fields = o.getClass().getDeclaredFields();
        for (Field field : fields) {
            boolean exist = field.isAnnotationPresent(FieldNotNull.class);
            if (!exist) {
                continue;
            }
            Object value = null;
            try {
                Method method = o.getClass().getMethod(BeanToMap.getMethodName(field.getName()));
                value = method.invoke(o);
            } catch (NoSuchMethodException e) {
                log.warn("反射方法异常！", e);
            } catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
                log.warn("反射参数异常！", e);
            }
            if (value == null || StringUtils.isEmpty(value.toString())) {
                return true;
            }
        }
        return false;
    }
}
