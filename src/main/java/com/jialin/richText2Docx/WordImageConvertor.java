package com.jialin.richText2Docx;

import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
import java.util.List;
import java.util.UUID;

import javax.imageio.ImageIO;

import org.apache.commons.codec.binary.Base64;

import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
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
	 * 添加图片到wordbao。
	 * @param wordMLPackage
	 * @param iswebsrc
	 * @param imageSrcs
	 * @throws Exception
	 */
	public static void addImgToPkg(WordprocessingMLPackage  wordMLPackage, boolean iswebsrc, List<String> imageSrcs) throws Exception {
		for (String imageSrc:imageSrcs
			 ) {
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
		byte[] bytes = convertImageToByteArray(file);
		addImageToPackage(wordMLPackage, bytes);
		}
	}


	/***
	 * 生成图片信息的标签块。
	 * @return
	 */
	public static String toDocxBodyBlock(
			String docFileName,
			String rid){

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
		sb1.append("<a:blip r:embed=\""+rid+"\" cstate=\"print\">");
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
	 *  Docx4j拥有一个由字节数组创建图片部件的工具方法, 随后将其添加到给定的包中. 为了能将图片添加
	 *  到一个段落中, 我们需要将图片转换成内联对象. 这也有一个方法, 方法需要文件名提示, 替换文本,
	 *  两个id标识符和一个是嵌入还是链接到的指示作为参数.
	 *  一个id用于文档中绘图对象不可见的属性, 另一个id用于图片本身不可见的绘制属性. 最后我们将内联
	 *  对象添加到段落中并将段落添加到包的主文档部件.
	 *
	 *  @param wordMLPackage 要添加图片的包
	 *  @param bytes         图片对应的字节数组
	 *  @throws Exception    不幸的createImageInline方法抛出一个异常(没有更多具体的异常类型)
	 */
	private static void addImageToPackage(WordprocessingMLPackage wordMLPackage,
										  byte[] bytes) throws Exception {
		BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

		int docPrId = 1;
		int cNvPrId = 2;
		Inline inline = imagePart.createImageInline("Filename hint","Alternative text", docPrId, cNvPrId, false);

		P paragraph = addInlineImageToParagraph(inline);

		wordMLPackage.getMainDocumentPart().addObject(paragraph);
	}

	/**
	 *  创建一个对象工厂并用它创建一个段落和一个可运行块R.
	 *  然后将可运行块添加到段落中. 接下来创建一个图画并将其添加到可运行块R中. 最后我们将内联
	 *  对象添加到图画中并返回段落对象.
	 *
	 * @param   inline 包含图片的内联对象.
	 * @return  包含图片的段落
	 */
	private static P addInlineImageToParagraph(Inline inline) {
		// 添加内联对象到一个段落中
		ObjectFactory factory = new ObjectFactory();
		P paragraph = factory.createP();
		R run = factory.createR();
		paragraph.getContent().add(run);
		Drawing drawing = factory.createDrawing();
		run.getContent().add(drawing);
		drawing.getAnchorOrInline().add(inline);
		return paragraph;
	}


	/**
	 * 将图片从文件转换成字节数组.
	 *
	 * @param file
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	private static byte[] convertImageToByteArray(File file) throws FileNotFoundException, IOException {
		InputStream is = new FileInputStream(file );
		long length = file.length();
		// You cannot create an array using a long, it needs to be an int.
		if (length > Integer.MAX_VALUE) {
			System.out.println("File too large!!");
		}
		byte[] bytes = new byte[(int)length];
		int offset = 0;
		int numRead = 0;
		while (offset < bytes.length && (numRead=is.read(bytes, offset, bytes.length-offset)) >= 0) {
			offset += numRead;
		}
		// Ensure all the bytes have been read
		if (offset < bytes.length) {
			System.out.println("Could not completely read file "+file.getName());
		}
		is.close();
		return bytes;
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
