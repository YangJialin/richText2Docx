package com.jialin.richText2Docx;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;
import org.jsoup.safety.Whitelist;
import org.jsoup.select.Elements;

import javax.sql.rowset.serial.SerialBlob;
import javax.sql.rowset.serial.SerialException;
import java.io.*;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;


public class RichText2Docx {
    String serverPath = "";
    private List<String> imgPath = new ArrayList<String>();

    Integer relationId = 6;
    List relationList = new ArrayList();
    String imageName ="";
    /**
     * 转化html为word并保存。
     *
     * @param data html
     * @param jurl 保存地址
     */
    public String resolveHtml(String data, String jurl) throws InvalidFormatException {

        Template template = getTemplate();
        if (template != null) {
            Map<String, Object> dataMap = new HashMap<String, Object>();
//            RichHtmlHandler handler = new RichHtmlHandler(data, appRootPath + File.separator);
//            data = handler.getHandledDocBodyBlock();
//            handledBase64Block += handler.getData(handler.getDocBase64BlockResults());
//            xmlimaHref += handler.getData(handler.getXmlImgRefs());

            Document document = Jsoup.parse(data);
            StringBuilder xmlData = new StringBuilder();
            //补全html标签。
            document.outputSettings().syntax(Document.OutputSettings.Syntax.html);
            document = Jsoup.parse(document.html());
            Elements elements = document.getAllElements();
            for (Element e : elements
            ) {
                switch (e.tagName()) {
                    case "p":
                        //一般P标签
                        if (!e.parent().tagName().equals("td") && !e.parent().tagName().equals("th")){
                            String pStr = e.text();
                            xmlData.append(handleElement("p",Entities.escape(pStr)));
                        }
                        break;

                    case "h1":
                        String h1Str = e.text();
                        xmlData.append(handleElement("h1",Entities.escape(h1Str)));
                        break;

                    case "h2":
                        String h2Str = e.text();
                        xmlData.append(handleElement("h2",Entities.escape(h2Str)));
                        break;

                    case "img":
                        String src = e.attr("src");
                        //图片绝对路径。
                        String srcRealPath = serverPath + src;
                        //收集图片，统一处理。
                        imgPath.add(srcRealPath);
                        break;

                    case "table":
                        Elements trs = e.getElementsByTag("tr");
                        xmlData.append(WordTableConvertor.handleTable(trs));
                        break;
                }
            }
            //拼接word的xml格式字符串。
            //1、获取富文本中的p标签及内容。
            //2、提取纯文本，拼接word到p标签。
            //3、循环执行，处理全部p标签。
            //4、将拼接好到写入模板。


            dataMap.put("content",xmlData.toString() );
            Writer wb = null;
            try {
                File xmlFile = new File(jurl + ".xml");
                OutputStream outTodisk = new FileOutputStream(xmlFile);

                ByteArrayOutputStream os = new ByteArrayOutputStream();
                wb = new BufferedWriter(new OutputStreamWriter(os, "UTF-8"));
                template.process(dataMap, wb);//写数据到模板
                wb.close();
                os.writeTo(outTodisk);
                outTodisk.flush();
                os.close();

                WordprocessingMLPackage wmlPackage = (WordprocessingMLPackage) WordprocessingMLPackage.load(new FileInputStream(xmlFile));
                if (imgPath.size()>0){
                    try {
                        WordImageConvertor.addImgToPkg(wmlPackage,true,imgPath);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }


                File file = new File(jurl + ".docx");
                wmlPackage.save(file, Docx4J.FLAG_SAVE_ZIP_FILE);

            } catch (FileNotFoundException e1) {
                e1.printStackTrace();
            } catch (TemplateException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (Docx4JException e) {
                e.printStackTrace();
            }

        }
        return "";
    }

    /**
     * 获取下载模板 .ftl文件
     *
     * @return
     * @createUser yangjialin
     * @createDate 2017年10月20日
     */
    public Template getTemplate() {
        //设置模本装置方法和路径,FreeMarker支持多种模板装载方法。可以重servlet，classpath，数据库装载，
        //这里我们的模板是放在resources/template 目录下面
        Configuration configuration = new Configuration();
        configuration.setDefaultEncoding("UTF-8");
        configuration.setClassForTemplateLoading(this.getClass(), "/");
        Template template = null;
        try {
            //blogTemplate.ftl为要装载的模板
            template = configuration.getTemplate("ftl/template.ftl", "UTF-8");
            return template;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取或生成本地路径，返回路径加文件名，不带扩展名。
     *
     * @param editDate
     * @param id
     * @return
     */
    public String getFileSavePath(String pageOfficeSavePath, Date editDate, String id) {
        StringBuilder zzPath = new StringBuilder(64);
        DateFormat df = new SimpleDateFormat("yyyy,MM,dd");
        String dstr;
        if (editDate != null) {
            dstr = df.format(editDate);
        } else {
            dstr = df.format(new Date());
        }

        String[] dateArray = dstr.split(",");

        zzPath.append(dateArray[0]);
        zzPath.append(File.separator);
        zzPath.append(dateArray[1]);
        zzPath.append(File.separator);
        zzPath.append(dateArray[2]);
        zzPath.append(File.separator);

        String path = zzPath.toString();
//        String pageOfficeSavePath = sysService.getProperty(KmisLocalConstant.PAGE_OFFICE_SAVE_PATH);
        File zzDir = new File(pageOfficeSavePath + File.separator + path);
        if (!zzDir.exists()) {
            zzDir.mkdirs();
        }
        return path + id;
    }

    /**
     * 将文件保存到字节数组中。
     *
     * @param is
     * @return
     * @throws IOException
     */
    private byte[] inputStreamToByte(InputStream is) throws IOException {
        ByteArrayOutputStream bAOutputStream = new ByteArrayOutputStream();
        int ch;
        while ((ch = is.read()) != -1) {
            bAOutputStream.write(ch);
        }
        byte data[] = bAOutputStream.toByteArray();
        bAOutputStream.close();
        return data;
    }

    /**
     * 处理文字为word xml标签。
     * @param elementType
     * @param text
     * @return
     */
    public String handleElement(String elementType, String text) {
        String h1TagStart = " <w:p w:rsidR=\"00036018\" w:rsidRDefault=\"00036018\" w:rsidP=\"00036018\">\n<w:pPr>\n" +
                "<w:pStyle w:val=\"1\"/>\n<w:rPr>\n<w:rFonts w:hint=\"eastAsia\"/>\n</w:rPr>\n</w:pPr>\n<w:r>\n<w:t>";
        String h1TagEnd = "</w:t>\n</w:r>\n</w:p>";

        String h2TagStart = "<w:p w:rsidR=\"00036018\" w:rsidRDefault=\"00036018\" w:rsidP=\"00036018\">\n<w:pPr>\n" +
                "<w:pStyle w:val=\"2\"/>\n</w:pPr>\n<w:r>\n<w:t>";
        String h2TagEnd = "</w:t>\n</w:r>\n</w:p>";

        String pTagStart = "<w:p w:rsidR=\"00036018\" w:rsidRDefault=\"00036018\">\n<w:r>\n<w:rPr>\n" +
                "<w:rFonts w:hint=\"eastAsia\"/>\n</w:rPr>\n<w:t>";
        String pTagEnd = "</w:t>\n</w:r>\n</w:p>";

        StringBuilder sb = new StringBuilder();

        switch (elementType) {
            case "h1":
                sb.append(h1TagStart+text+h1TagEnd);
                break;
            case "h2":
                sb.append(h2TagStart+text+h2TagEnd);
                break;
            case "p":
                sb.append(pTagStart+text+pTagEnd);
                break;
        }
        return sb.toString();
    }

    public String getData(List<String> list){
        String data = "";
        if (list != null && list.size() > 0) {
            for (String string : list) {
                data += string + "\n";
            }
        }
        return data;
    }
}
