package com.jialin.richText2Docx;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
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
    private List<String> docBase64BlockResults = new ArrayList<String>();

    /**
     * 转化html为word并保存。
     *
     * @param data html
     * @param jurl 保存地址
     */
    public String resolveHtml(String data, String jurl) {
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
                        String pStr = e.text();
                        xmlData.append(handleElement("p",pStr));
                        break;
                    //表格中的P标签。

                    case "h1":
                        String h1Str = e.text();
                        xmlData.append(handleElement("h1",h1Str));
                        break;

                    case "h2":
                        String h2Str = e.text();
                        xmlData.append(handleElement("h2",h2Str));
                        break;

                    case "img":
                        //图片。TODO
                        String src = e.attr("src");
                        //图片绝对路径。
                        String srcRealPath = serverPath + src;
                        File imageFile = new File(srcRealPath);
                        //获取文件名。
                        String imageFileName = imageFile.getName();
                        //获取图片扩展名。
                        String fileTypeName = WordImageConvertor.getFileSuffix(srcRealPath);
                        //图片的新名字。
                        String uuid = UUID.randomUUID().toString();
                        String docFileName = "image" + uuid + "."+ fileTypeName;

                        // 得到文件的word xml的代码块。
                        String handledDocxBodyBlock = WordImageConvertor.toDocxBodyBlock(docFileName,uuid);

                        //将图片转换成base64编码的字符串
                        String base64Content = "";
                        try {
                            base64Content = WordImageConvertor.imageToBase64(true,srcRealPath);
                        } catch (IOException ex) {
                            ex.printStackTrace();
                        } catch (Exception ex) {
                            ex.printStackTrace();
                        }
                        //生成图片的base4块.
                        String docBase64BlockResult = WordImageConvertor.generateImageBase64Block(docFileName,
                                fileTypeName, base64Content);
                        docBase64BlockResults.add(docBase64BlockResult);

                        xmlData.append(handledDocxBodyBlock);
                        break;
                }
            }
            //拼接word的xml格式字符串。
            //1、获取富文本中的p标签及内容。
            //2、提取纯文本，拼接word到p标签。
            //3、循环执行，处理全部p标签。
            //4、将拼接好到写入模板。


            dataMap.put("content",xmlData.toString() );
            dataMap.put("docBase64BlockResults",getData(docBase64BlockResults));
            if (docBase64BlockResults.size()>0){
                dataMap.put("relationship","<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.jpeg\"/>");
            }else{
                dataMap.put("relationship","");
            }
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

        String imgTagStart ="<w:p w:rsidR=\"00036018\" w:rsidRDefault=\"00036018\">\n" +
                "                        <w:r>\n" +
                "                            <w:rPr>\n" +
                "                                <w:rFonts w:hint=\"eastAsia\"/>\n" +
                "                                <w:noProof/>\n" +
                "                            </w:rPr>\n" +
                "                            <w:drawing>\n" +
                "                                <wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">\n" +
                "                                    <wp:extent cx=\"5270500\" cy=\"5340985\"/>\n" +
                "                                    <wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"5715\"/>\n" +
                "                                    <wp:docPr id=\"1\" name=\"图片 1\" descr=\"\"/>\n" +
                "                                    <wp:cNvGraphicFramePr>\n" +
                "                                        <a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/>\n" +
                "                                    </wp:cNvGraphicFramePr>\n" +
                "                                    <a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n" +
                "                                        <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n" +
                "                                            <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n" +
                "                                                <pic:nvPicPr>\n" +
                "                                                    <pic:cNvPr id=\"1\" name=\"name.jpg\"/>\n" +
                "                                                    <pic:cNvPicPr/>\n" +
                "                                                </pic:nvPicPr>\n" +
                "                                                <pic:blipFill>\n" +
                "                                                    <a:blip r:embed=\"rId4\" cstate=\"print\">\n" +
                "                                                        <a:extLst>\n" +
                "                                                            <a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\">\n" +
                "                                                                <a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/>\n" +
                "                                                            </a:ext>\n" +
                "                                                        </a:extLst>\n" +
                "                                                    </a:blip>\n" +
                "                                                    <a:stretch>\n" +
                "                                                        <a:fillRect/>\n" +
                "                                                    </a:stretch>\n" +
                "                                                </pic:blipFill>\n" +
                "                                                <pic:spPr>\n" +
                "                                                    <a:xfrm>\n" +
                "                                                        <a:off x=\"0\" y=\"0\"/>\n" +
                "                                                        <a:ext cx=\"5270500\" cy=\"5340985\"/>\n" +
                "                                                    </a:xfrm>\n" +
                "                                                    <a:prstGeom prst=\"rect\">\n" +
                "                                                        <a:avLst/>\n" +
                "                                                    </a:prstGeom>\n" +
                "                                                </pic:spPr>\n" +
                "                                            </pic:pic>\n" +
                "                                        </a:graphicData>\n" +
                "                                    </a:graphic>\n" +
                "                                </wp:inline>\n" +
                "                            </w:drawing>\n" +
                "                        </w:r>\n" +
                "                    </w:p>";
        String imgTagEnd = "";

        String tableTagStart="<w:tbl>\n" +
                "                        <w:tblPr>\n" +
                "                            <w:tblStyle w:val=\"a5\"/>\n" +
                "                            <w:tblW w:w=\"0\" w:type=\"auto\"/>\n" +
                "                            <w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/>\n" +
                "                        </w:tblPr>\n" +
                "                        <w:tblGrid>\n" +
                "                            <w:gridCol w:w=\"2763\"/>\n" +
                "                            <w:gridCol w:w=\"2763\"/>\n" +
                "                            <w:gridCol w:w=\"2764\"/>\n" +
                "                        </w:tblGrid>\n" +
                "                        <w:tr w:rsidR=\"00036018\" w:rsidTr=\"00036018\">\n" +
                "                            <w:tc>\n" +
                "                                <w:tcPr>\n" +
                "                                    <w:tcW w:w=\"2763\" w:type=\"dxa\"/>\n" +
                "                                </w:tcPr>\n" +
                "                                <w:p w:rsidR=\"00036018\" w:rsidRDefault=\"00036018\" w:rsidP=\"00036018\">\n" +
                "                                    <w:pPr>\n" +
                "                                        <w:rPr>\n" +
                "                                            <w:rFonts w:hint=\"eastAsia\"/>\n" +
                "                                        </w:rPr>\n" +
                "                                    </w:pPr>\n" +
                "                                    <w:r>\n" +
                "                                        <w:rPr>\n" +
                "                                            <w:rFonts w:hint=\"eastAsia\"/>\n" +
                "                                        </w:rPr>\n" +
                "                                        <w:t>表格第一行第一列</w:t>\n" +
                "                                    </w:r>\n" +
                "                                </w:p>\n" +
                "                            </w:tc>\n" +
                "                        </w:tr>\n" +
                "                    </w:tbl>";
        String tableTagEnd="";

        String imgDataTagStart = "<pkg:part pkg:name=\"/word/media/image1.jpeg\" pkg:contentType=\"image/jpeg\" pkg:compression=\"store\">\n" +
                "        <pkg:binaryData>data</pkg:binaryData>\n" +
                "    </pkg:part>";
        String imgDataTagEnd = "</pkg:binaryData>\n" +
        "    </pkg:part>";
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
            case "table":
                //TODO
                break;
            case "img":
                //TODO
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
