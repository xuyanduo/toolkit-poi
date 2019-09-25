package service;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class BeanToMap<T> {

    public static String getMethodName(String fieldName) {
        byte[] buffer = fieldName.getBytes();
        buffer[0] = (byte) (buffer[0] - 32);
        String name = new String(buffer);
        return "get" + name;
    }

    /**
     * @Description  将实体转换成map
     * @Param [entity] 实体
     * @return  map
     **/
    public Map<String, Object> getMap(T entity) {
        if (entity instanceof Map) {
            return (Map<String, Object>) entity;
        }
        Field[] fields = entity.getClass().getDeclaredFields();
        Map<String, Object> map = new HashMap<>();
        for (int j = 0; j < fields.length; j++) {
            try {
                Method method = entity.getClass().getMethod(getMethodName(fields[j].getName()));
                map.put(fields[j].getName(), method.invoke(entity));
            } catch (Exception e) {
                log.info("bean2Map 反射异常", e);
            }
        }
        return map;
    }

    /**
     * @Description  bean 转换成map
     * @Param [list] list集合
     * @return  反馈转换后的集合
     **/
    public List<Map<String, Object>> changeBean2Map(List<T> list) {
        List<Map<String, Object>> result = new ArrayList<>();
        if (CollectionUtils.isEmpty(list)) {
            return result;
        }
        for (T t : list) {
            Map<String, Object> hm = getMap(t);
            result.add(hm);
        }
        return result;
    }
}

