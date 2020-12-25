package cn.iscgm.poi.utils;

import net.coobird.thumbnailator.Thumbnails;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
import java.nio.file.Files;

/**
 * 压缩图片工具类
 *
 * @author cgm
 */
public class ThumbnailsUtils {
    public static void main(String[] args) {
        //存放照片的路径
        String parentPath = "D:\\新桌面\\照片";
        File dir = new File(parentPath);
        for (File pic : dir.listFiles()) {
            if (pic.isDirectory()) {
                continue;
            }
            if (pic.getName().endsWith(".jpg")) {
                try {
                    byte[] bytes = compressPicCycle(Files.readAllBytes(pic.toPath()), 300, 0.5);
                    BufferedOutputStream outputStream = new BufferedOutputStream(new FileOutputStream(pic));
                    outputStream.write(bytes);
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }


    /**
     * @param bytes       原图片字节数组
     * @param desFileSize 指定图片大小,单位 kb
     * @param accuracy    精度,递归压缩的比率,建议小于0.9
     * @return
     */
    private static byte[] compressPicCycle(byte[] bytes, long desFileSize, double accuracy) throws IOException {
        long fileSize = bytes.length;
        System.out.println("=====fileSize======== " + fileSize);
        // 判断图片大小是否小于指定图片大小
        if (fileSize <= desFileSize * 1024) {
            return bytes;
        }
        //计算宽高
        BufferedImage bim = ImageIO.read(new ByteArrayInputStream(bytes));
        int imgWidth = bim.getWidth();
        System.out.println(imgWidth + "====imgWidth=====");
        int imgHeight = bim.getHeight();
        int desWidth = new BigDecimal(imgWidth).multiply(new BigDecimal(accuracy)).intValue();
        System.out.println(desWidth + "====desWidth=====");
        int desHeight = new BigDecimal(imgHeight).multiply(new BigDecimal(accuracy)).intValue();
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        //字节输出流（写入到内存）
        Thumbnails.of(new ByteArrayInputStream(bytes)).size(desWidth, desHeight).outputQuality(accuracy).toOutputStream(byteArrayOutputStream);
        //如果不满足要求,递归直至满足要求
        return compressPicCycle(byteArrayOutputStream.toByteArray(), desFileSize, accuracy);
    }

}
