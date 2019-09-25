package utils;


import bean.ExcelData;
import com.alibaba.excel.util.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.formula.functions.T;
import service.BeanToMap;
import utils.DateFormateUtils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URL;
import java.net.URLDecoder;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.UUID;

/**
 * @author xyd
 * @descripe excel 导出
 */
@Slf4j
public class PoiExcelExport {

    public static void exportExcel(HttpServletRequest request, HttpServletResponse response, ExcelData data) {
        log.info("导出解析开始，fileName:{}", data.getFileName());
        try {
            //实例化HSSFWorkbook
            HSSFWorkbook workbook = fillExcelWorkbook(data);
            //设置浏览器下载
            setBrowser(request, response, workbook, data.getFileName());
            log.info("导出解析成功!");
        } catch (Exception e) {
            log.info("导出解析失败!", e);
        }
    }

    /**
     * @return 字节数组
     * @Description 返回导出excel的字节流
     * @Param [data] exceldata
     **/
    public static byte[] exportExcel(ExcelData data, String type) {
        log.info("导出解析开始，fileName:{}", data.getFileName());
        //实例化HSSFWorkbook
        HSSFWorkbook workbook = fillExcelWorkbook(data);
        log.info("导出解析成功!");
        String filePath = saveExcelFile(workbook, type);
        if (StringUtils.isEmpty(filePath)) {
            return null;
        }
        return FileDownLoadUtil.download(filePath);
    }


    /**
     * 保存临时文件
     *
     * @param workbook 表格数据
     * @return 路径
     */
    public static String saveExcelFile(HSSFWorkbook workbook, String type) {
        String rootPath="";
        if(File.separator.equals("\\")){
            rootPath = "C:";
        }
        log.info("根路径：{}", rootPath);
        StringBuilder sb = new StringBuilder();
        sb.append(rootPath).append(File.separator).append("download").append(File.separator).append(type);
        checkPathExist(sb.toString());
        sb.append(File.separator).append("temp-").append(System.currentTimeMillis()).append("-").append(UUID.randomUUID()).append(
                ".xls");
        log.info("文件路径：{}", sb.toString());
        try {
            FileOutputStream output = new FileOutputStream(sb.toString());
            workbook.write(output);
        } catch (IOException e) {
            log.warn("生成excel文件异常！！！");
            return null;
        }
        return sb.toString();
    }

    /**
     * 判断文件夹是否存在，如果不存在则新建
     *
     * @param dirPath 文件夹路径
     */
    private static void checkPathExist(String dirPath) {
        File file = new File(dirPath);
        if (!file.exists()) {
            file.mkdirs();
        }
    }

    /**
     * @return workbook
     * @Description 填充excel
     * @Param [data] exceldata
     **/
    private static HSSFWorkbook fillExcelWorkbook(ExcelData data) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建一个Excel表单，参数为sheet的名字
        HSSFSheet sheet = workbook.createSheet("sheet");
        //设置表头
        setTitle(workbook, sheet, data.getHeads());
        //设置单元格并赋值
        setData(sheet, data.getList(), data.getCols());
        return workbook;
    }

    /**
     * 方法名：setTitle
     * 功能：设置表头
     */
    private static void setTitle(HSSFWorkbook workbook, HSSFSheet sheet, String[] str) {
        try {
            HSSFRow row = sheet.createRow(0);
            //设置列宽，setColumnWidth的第二个参数要乘以256，这个参数的单位是1/256个字符宽度
            for (int i = 0; i <= str.length; i++) {
                sheet.setColumnWidth(i, 15 * 256);
            }
            //设置为居中加粗,格式化时间格式
            HSSFCellStyle style = workbook.createCellStyle();
            HSSFFont font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            style.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
            //创建表头名称
            HSSFCell cell;
            for (int j = 0; j < str.length; j++) {
                cell = row.createCell(j);
                cell.setCellValue(str[j]);
                cell.setCellStyle(style);
            }
        } catch (Exception e) {
            log.info("导出时设置表头失败！", e);
        }
    }

    /**
     * 方法名：setData
     * 功能：表格赋值
     */
    private static void setData(HSSFSheet sheet, List<T> list, String[] cols) {
        try {
            BeanToMap<T> btm = new BeanToMap<>();
            for (int rowNum = 1; rowNum <= list.size(); rowNum++) {
                Map<String, Object> hm = btm.getMap(list.get(rowNum - 1));
                HSSFRow row = sheet.createRow(rowNum);
                // 读取数据值
                for (int k = 0; k < cols.length; k++) {
                    HSSFCell hssfcell = row.createCell(k);
                    Object obj = hm.get(cols[k]);
                    if (null == obj) {
                        hssfcell.setCellValue("");
                        continue;
                    }
                    if (obj instanceof Date) {
                        hssfcell.setCellValue(DateFormateUtils.date2String((Date) obj, "yyyy-MM-dd HH:mm:ss"));
                        continue;
                    }
                    hssfcell.setCellValue(obj.toString());
                }
            }
            log.info("表格赋值成功！");
        } catch (Exception e) {
            log.info("表格赋值失败！", e);
        }
    }

    /**
     * 方法名：setBrowser
     * 功能：使用浏览器下载
     */
    private static void setBrowser(HttpServletRequest request, HttpServletResponse response, HSSFWorkbook workbook,
                                   String fileName) {
        try {
            response.reset();
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + new String((fileName + ".xls").getBytes(), "iso-8859-1"));
            OutputStream os = new BufferedOutputStream(response.getOutputStream());
            //将excel写入到输出流中
            workbook.write(os);
            os.flush();
            os.close();
            log.info("设置浏览器下载成功！");
        } catch (Exception e) {
            log.info("设置浏览器下载失败！", e);
        }
    }

}
