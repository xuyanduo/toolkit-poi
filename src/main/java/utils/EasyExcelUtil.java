package utils;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.StringUtils;
import com.alibaba.excel.event.adapter.DefaultAnalysisEventListenerAdapter;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.util.List;

/**
 * @ClassName EasyExcelUtil
 * @Description easyExcel操作excel工具类
 * @Auther William
 * @Date 2019/6/5 11:35
 * @Version 1.0
 */
@Slf4j
public class EasyExcelUtil {


    /*
     *@Description: 读取Excel文件内容
     *@param in excel文件流
     *@param tClass 对应excel实体bean
     *@return: 对应excel实体bean的list
     *@Author:  William
     *@Date:  2019/6/5 13:24
     */
    public static <T> List<T> getExcelContent(InputStream in, Class<T> tClass, String fileName) {
        AnalysisEventListener<T> listener = new DefaultAnalysisEventListenerAdapter<>();
        return getExcelContent(in, tClass, fileName, listener, 1, 1);
    }


    /**
     * @param in       输入流
     * @param tClass   类信息
     * @param fileName 文件名称
     * @param listener 监听器  实现 AnalysisEventListener抽象方法 提供了默认的DefaultAnalysisEventListenerAdapter
     * @param <T>      数据结果类型
     * @param sheetNo  sheet页  从1开始
     * @param headLineNum 列头的行数量
     * @return 数据结果
     * @Description: 获取excel content 可传递自定义的监听listener
     */
    public static <T> List<T> getExcelContent(InputStream in, Class<T> tClass, String fileName,
                                                AnalysisEventListener<T> listener, int sheetNo, int headLineNum) {
        try {
            ExcelReader excelReader = new ExcelReader(in, getExcelTypeEnum(fileName), null, listener);
            // 第二个参数为表头行数，按照实际设置
            excelReader.read(new Sheet(sheetNo, headLineNum, (Class<? extends BaseRowModel>) tClass));
            listener.getData();
        } catch (Exception e) {
            log.warn("读取excel 异常", e);
        }
        return listener.getData();
    }


    /**
     * @Description  有默认的监听器 需要传入sheet页 头数量
     * @Param [in, tClass, fileName, sheetNo, headLineNum] 文件流,类信息,文件名称,sheet第几页从1开始,列头数量
     * @return 数据的集合
     **/
    public static <T> List<T> getExcelContent(InputStream in, Class<T> tClass, String fileName, int sheetNo, int headLineNum) {
        AnalysisEventListener<T> listener = new DefaultAnalysisEventListenerAdapter<>();
        getExcelContent(in, tClass, fileName, listener, sheetNo, headLineNum);
        return listener.getData();
    }


    /**
     * @param fileName 文件名称
     * @return excel类型枚举
     * @throws Exception 异常
     */
    private static ExcelTypeEnum getExcelTypeEnum(String fileName) throws Exception {
        if (StringUtils.isEmpty(fileName)) {
            throw new Exception("空文件名！");
        }
        if (fileName.endsWith(".xls")) {
            return ExcelTypeEnum.XLS;
        } else {
            return ExcelTypeEnum.XLSX;
        }
    }
}