package utils;

import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.*;

/**
 * @Auther: xyd
 * @Date: 2019/7/16 08:52
 * @Description:
 */
@Slf4j
public class FileDownLoadUtil {

    public static byte[] download(String path) {
        try {
            log.info("文件下载的路径：{}", path);
            // path是指欲下载的文件的路径。
            File file = new File(path);
            // 取得文件名。
            String filename = file.getName();
            log.info("文件名称：{}", filename);

            // 以流的形式下载文件。
            InputStream fis = new BufferedInputStream(new FileInputStream(path));
            byte[] buffer = new byte[fis.available()];
            fis.read(buffer);
            return buffer;
        } catch (IOException ex) {
            log.warn("下载文件异常！", ex);
            return null;
        }
    }


    /**
     * @Description  下载编译后相对路径下的文件
     * @Param [relativePath] 相对路径
     * @return 返回文件字节流
     **/
    public static byte[] downloadJarFile(String relativePath) {
        log.info("下载class中模板文件的相对路径：{}", relativePath);
        Resource res = new ClassPathResource(relativePath);
        try {
            InputStream fis = res.getInputStream();
            byte[] buffer = new byte[fis.available()];
            fis.read(buffer);
            return buffer;
        } catch (IOException e) {
            log.warn("下载文件异常！", e);
            return null;
        }
    }
}
