package bean;

import lombok.Data;

import java.util.List;

/**
 * @Auther: xyd
 * @Date: 2019/5/9 14:08
 * @Description:
 */

@Data
public class ExcelData<T> {
    /**
     * 文件名称
     **/
    private String fileName;
    /**
     * 表头
     **/
    private String[] heads;
    /**
     * 列
     **/
    private String[] cols;
    /**
     * 数据集合
     **/
    private List<T> list;

    public ExcelData(String fileName, String[] heads, String[] cols, List<T> list){
        this.fileName = fileName;
        this.heads = heads;
        this.cols = cols;
        this.list = list;
    }
}
