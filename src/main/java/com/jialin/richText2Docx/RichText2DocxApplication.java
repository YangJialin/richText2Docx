package com.jialin.richText2Docx;

import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class RichText2DocxApplication {

	public static void main(String[] args) {
		SpringApplication.run(RichText2DocxApplication.class, args);
		RichText2Docx text2Docx = new RichText2Docx();
		String data = "<h1>\n" +
				"    标题1<br/>\n" +
				"</h1>\n" +
				"<h2>\n" +
				"    标题2<br/>\n" +
				"</h2>\n" +
				"<p>\n" +
				"    文本内容<br/>\n" +
				"</p>\n" +
				"<p>\n" +
				"    <img src=\"https://gss1.bdstatic.com/-vo3dSag_xI4khGkpoWK1HF6hhy/baike/c0%3Dbaike272%2C5%2C5%2C272%2C90/sign=22139b28761ed21b6dc426b7cc07b6a1/ac4bd11373f08202ff289be540fbfbedaa641bb5.jpg\" width=\"521\" height=\"356\"/>\n" +
				"</p>\n" +
				"<table>\n" +
				"    <tbody>\n" +
				"        <tr class=\"firstRow\">\n" +
				"            <td width=\"260\" valign=\"top\" style=\"word-break: break-all;\">\n" +
				"                <p>1</p>\n" +
				"            </td>\n" +
				"            <td width=\"260\" valign=\"top\" style=\"word-break: break-all;\">\n" +
				"                <p>2</p>\n" +
				"            </td>\n" +
				"            <td width=\"260\" valign=\"top\" style=\"word-break: break-all;\">\n" +
				"                <p>3</p>\n" +
				"            </td>\n" +
				"        </tr>\n" +
				"        <tr>\n" +
				"            <td width=\"260\" valign=\"top\" style=\"word-break: break-all;\">\n" +
				"                <p>4</p>\n" +
				"            </td>\n" +
				"            <td width=\"260\" valign=\"top\" style=\"word-break: break-all;\">\n" +
				"                <p>5</p>\n" +
				"            </td>\n" +
				"            <td width=\"260\" valign=\"top\" style=\"word-break: break-all;\">\n" +
				"                <p>6</p>\n" +
				"            </td>\n" +
				"        </tr>\n" +
				"    </tbody>\n" +
				"</table>\n" +
				"<p>\n" +
				"    <br/>\n" +
				"</p>";
		String path = "/Users/gallin/Workspace/kmis/pageoffice/test";
		try {
			text2Docx.resolveHtml(data, path);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
	}

}
