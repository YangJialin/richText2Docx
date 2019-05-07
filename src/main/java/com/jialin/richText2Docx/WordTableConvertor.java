package com.jialin.richText2Docx;

import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;
import org.jsoup.select.Elements;

/**
 * @author yangjialin
 * @Description:WORD 文档表格转换器
 */
public class WordTableConvertor {

    /***
     * 处理表格富文本。
     * @return
     */
    public static String handleTable(Elements trs) {

        StringBuilder tableXML = new StringBuilder();
        tableXML.append("<w:tbl><w:tblPr><w:tblStyle w:val=\"a5\"/><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/>\n" +
                "</w:tblPr><w:tblGrid><w:gridCol w:w=\"2763\"/><w:gridCol w:w=\"2763\"/><w:gridCol w:w=\"2764\"/></w:tblGrid>\n");
        for (Element e : trs
        ) {

            tableXML.append("<w:tr w:rsidR=\"00036018\" w:rsidTr=\"00036018\">\n");
            Elements subElements = e.getAllElements();
            for (Element subE : subElements
            ) {
                switch (subE.tagName()) {
                    case "td":
                        tableXML.append("<w:tc><w:tcPr><w:tcW w:w=\"2763\" w:type=\"dxa\"/>\n" +
                                "</w:tcPr><w:p w:rsidR=\"00036018\" w:rsidRDefault=\"00036018\" w:rsidP=\"00036018\">\n" +
                                "<w:pPr><w:rPr><w:rFonts w:hint=\"eastAsia\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\"/>\n" +
                                "</w:rPr><w:t>");
                        tableXML.append(Entities.escape(subE.text()));
                        tableXML.append("</w:t></w:r></w:p></w:tc>\n");
                        break;

                    case "th":
                        tableXML.append("<w:tc><w:tcPr><w:tcW w:w=\"2763\" w:type=\"dxa\"/>\n" +
                                "</w:tcPr><w:p w:rsidR=\"00036018\" w:rsidRDefault=\"00036018\" w:rsidP=\"00036018\">\n" +
                                "<w:pPr><w:rPr><w:rFonts w:hint=\"eastAsia\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\"/>\n" +
                                "</w:rPr><w:t>");
                        tableXML.append(Entities.escape(subE.text()));
                        tableXML.append("</w:t></w:r></w:p></w:tc>\n");
                        break;
                }
            }
            tableXML.append("</w:tr>\n");

        }
        tableXML.append("</w:tbl>");

        return tableXML.toString();
    }
}
