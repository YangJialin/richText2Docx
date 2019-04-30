package com.jialin.richText2Docx;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.UUID;

public class ImageRequest {
    /**
     * @param urlStr
     * @return
     */
    public static File getWebImg(String urlStr) throws Exception {
        //new一个URL对象  
        URL url = new URL(urlStr);
        String localTempLocation="/Users/gallin/Workspace/kmis/pageoffice/test/img";
        File l = new File(localTempLocation);
        if (!l.exists()){
            l.mkdirs();
        }
        //打开链接  
        HttpURLConnection conn = (HttpURLConnection)url.openConnection();
        //设置请求方式为"GET"  
        conn.setRequestMethod("GET");
        //超时响应时间为5秒  
        conn.setConnectTimeout(5 * 1000);
        //通过输入流获取图片数据  
        InputStream inStream = conn.getInputStream();
        //得到图片的二进制数据，以二进制封装得到数据，具有通用性  
        byte[] data = readInputStream(inStream);

        int lastIndex = urlStr.lastIndexOf(".");
        String suffix = urlStr.substring(lastIndex + 1);

        File imageFile = new File(localTempLocation+File.separator+ UUID.randomUUID()+"."+suffix);
        //创建输出流  
        FileOutputStream outStream = new FileOutputStream(imageFile);
        //写入数据  
        outStream.write(data);
        //关闭输出流  
        outStream.close();
        return imageFile;
    }

    public static byte[] readInputStream(InputStream inStream) throws Exception{
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        //创建一个Buffer字符串  
        byte[] buffer = new byte[1024];
        //每次读取的字符串长度，如果为-1，代表全部读取完毕  
        int len = 0;
        //使用一个输入流从buffer里把数据读取出来  
        while( (len=inStream.read(buffer)) != -1 ){
            //用输出流往buffer里写入数据，中间参数代表从哪个位置开始读，len代表读取的长度  
            outStream.write(buffer, 0, len);
        }
        //关闭输入流  
        inStream.close();
        //把outStream里的数据写入内存  
        return outStream.toByteArray();
    }
}  