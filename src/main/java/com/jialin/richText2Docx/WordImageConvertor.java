package com.jialin.richText2Docx;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.UUID;

import javax.imageio.ImageIO;

import org.apache.commons.codec.binary.Base64;

import sun.misc.BASE64Encoder;



/**   
* @Description:WORD 文档图片转换器
 * @author yangjialin
*   
*/
public class WordImageConvertor {
	
	//private static Const WORD_IMAGE_SHAPE_TYPE_ID="";
	public static Integer imgIdCount = 1;
	/**   
	* @Description: 将图片转换成base64编码的字符串  
	* @param @param imageSrc 文件路径
	* @return String
	* @throws IOException
	* @throws
	*/ 
	public static String imageToBase64(boolean iswebsrc,String imageSrc) throws Exception {
		File file;
		if (iswebsrc){
			file = ImageRequest.getWebImg(imageSrc);
		}else {
			file=new File(imageSrc);
		}
		//判断文件是否存在
		if(file == null || !file.exists()){
			throw new FileNotFoundException("文件不存在！");
		}
		StringBuilder pictureBuffer = new StringBuilder();
		FileInputStream input=new FileInputStream(file);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
		//读取文件
		BASE64Encoder encoder=new BASE64Encoder();
		byte[] temp = new byte[1024];
        for(int len = input.read(temp); len != -1;len = input.read(temp)){
            out.write(temp, 0, len);
        }
        pictureBuffer.append(new String( Base64.encodeBase64Chunked(out.toByteArray())));
        input.close();
		return pictureBuffer.toString();
	}


	/***
	 * 生成图片信息的标签块。
	 * @return
	 */
	public static String toDocxBodyBlock(
			String docFileName,
			String uuid){

		StringBuilder sb1=new StringBuilder();
		
		sb1.append(" <w:p w:rsidR=\"00036018\" w:rsidRDefault=\"00036018\">\n<w:r>\n" +
				"<w:rPr>\n<w:rFonts w:hint=\"eastAsia\"/>\n<w:noProof/>\n</w:rPr>");
		sb1.append("<w:drawing>");
		sb1.append("<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
		sb1.append("<wp:extent cx=\"5270500\" cy=\"5340985\"/>");
		sb1.append("<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"5715\"/>");
		sb1.append("<wp:docPr id=\"" + imgIdCount+"\" name=\"图片 1\" descr=\" \"/>");
		sb1.append("<wp:cNvGraphicFramePr> <a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>");
		sb1.append("<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
		sb1.append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
		sb1.append("<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
		sb1.append("<pic:nvPicPr>");
		sb1.append("<pic:cNvPr id=\""+imgIdCount+"\" name=\"" + docFileName +"\"/><pic:cNvPicPr/>");
		sb1.append(" </pic:nvPicPr>");
		sb1.append("<pic:blipFill>");
		sb1.append("<a:blip r:embed=\"rId4\" cstate=\"print\">");
		sb1.append("<a:extLst>");
		sb1.append("<a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\">");
		sb1.append("<a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/>");
		sb1.append(" </a:ext>");
		sb1.append(" </a:extLst>");
		sb1.append(" </a:blip>");
		sb1.append(" <a:stretch>");
		sb1.append(" <a:fillRect/>");
		sb1.append("  </a:stretch>");
		sb1.append(" </pic:blipFill>");
		sb1.append("<pic:spPr>");
		sb1.append("  <a:xfrm>");
		sb1.append(" <a:off x=\"0\" y=\"0\"/>");
		sb1.append(" <a:ext cx=\"5270500\" cy=\"5340985\"/>");
		sb1.append(" </a:xfrm>");
		sb1.append("<a:prstGeom prst=\"rect\">\n" +
				"<a:avLst/>\n" +
				"</a:prstGeom>\n" +
				"</pic:spPr>\n" +
				"</pic:pic>\n" +
				"</a:graphicData>\n" +
				"</a:graphic>\n" +
				"</wp:inline>\n" +
				" </w:drawing>\n" +
				"</w:r>\n" +
				"</w:p>");
		imgIdCount++;
		return sb1.toString();
	}
	
	/**
	 * 生成图片的base4块
	 * @param fileTypeName
	 * @param base64Content
	 * @return
	 */
	public static String generateImageBase64Block(String imgLoacation,
									String fileTypeName,String base64Content){

		
		StringBuilder sb=new StringBuilder();
		sb.append("<pkg:part pkg:name=\"/word/media/"+imgLoacation+"\" pkg:contentType=\""+getImageContentType(fileTypeName)+"\" pkg:compression=\"store\">");
		sb.append(base64Content);
		sb.append("</pkg:part>");
		return sb.toString();
	}
	

	
	private static String getImageContentType(String fileTypeName){
		String result="image/jpeg";
		//http://tools.jb51.net/table/http_content_type
		if(fileTypeName.equals("tif") || fileTypeName.equals("tiff")){
			result="image/tiff";
		}else if(fileTypeName.equals("fax")){
			result="image/fax";
		}else if(fileTypeName.equals("gif")){
			result="image/gif";
		}else if(fileTypeName.equals("ico")){
			result="image/x-icon";
		}else if(fileTypeName.equals("jfif") || fileTypeName.equals("jpe") 
					||fileTypeName.equals("jpeg")  ||fileTypeName.equals("jpg")){
			result="image/jpeg";
		}else if(fileTypeName.equals("net")){
			result="image/pnetvue";
		}else if(fileTypeName.equals("png") || fileTypeName.equals("bmp") ){
			result="image/png";
		}else if(fileTypeName.equals("rp")){
			result="image/vnd.rn-realpix";
		}else if(fileTypeName.equals("rp")){
			result="image/vnd.rn-realpix";
		}
		
		return result;
	}
	
	/**
	 * 获取图片后缀名，如jpg,png等
	 * @param srcRealPath 图片绝对路径
	 * @return
	 */
	public static String getFileSuffix(String srcRealPath){
		int lastIndex = srcRealPath.lastIndexOf(".");
		String suffix = srcRealPath.substring(lastIndex + 1);
		return suffix;
	}
	
	
	
	
}
