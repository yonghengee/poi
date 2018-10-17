package com.yqh.bbct.excel;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Excel解析
 * <p>
 * Created by wyh in 2018/9/14 15:58
 **/
public class CommonUtils {
    /**
     * 从实体中解析出字段数据
     *
     * @param data  可能为pojo或者map 从field中解析
     * @param field 字段名称
     * @return
     */

    @SuppressWarnings("rawtypes")
    public static Object getValue(Map<String, Object> data, String field) {
        return data.get(field);
    }

    /**
     * 判断是否为数字
     *
     * @param v
     * @return
     */
    public static boolean isNumber(Object v) {

        if (v == null) return false;

        if (v instanceof Number) {
            return true;
        } else if (v.toString().matches("^\\d+$")) {
            return true;
        } else if (v.toString().matches("^-?\\d+\\.?\\d+$")) {
            return true;
        } else {
            try {
                Double.parseDouble(v.toString());
                return true;
            } catch (Exception e) {
                return false;
            }


        }

    }

    /**
     * 返回 #{} 或者 ${} 中包含的值
     *
     * @param str
     * @return eg:#{name} ${ages}
     */
    public static String[] getWildcard(String str) {

        List<String> list = new ArrayList<String>();

        int start = 0;
        while (start < str.length() && start >= 0) {

            start = str.indexOf("{", start);

            int end = str.indexOf("}", start);
            if (start > 0) {
                String wc = str.substring(start - 1, end + 1);

                list.add(wc);
            }

            if (start < 0) break;

            start = end + 1;

        }
        return list.toArray(new String[0]);
    }
}
