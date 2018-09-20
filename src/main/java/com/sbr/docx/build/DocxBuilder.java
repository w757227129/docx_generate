package com.sbr.docx.build;

import java.io.File;

import org.apache.log4j.Logger;

import com.sbr.docx.util.FreemarkerUtil;
import com.sbr.docx.util.DocxZipUtil;






/**
 * 
  * 生成docx文件的单例类
  *
  * @author 曹传红
  * @date 2018年7月29日 下午6:58:16
  * @version 1.0 
  *
 */
public class DocxBuilder {

	private static Logger logger = Logger.getLogger(DocxBuilder.class);
	
	
	private DocxBuilder(){}
	private FreemarkerUtil fu = new FreemarkerUtil();
	
	private static class DocxBuilderHodler{
		private static final DocxBuilder docxBuilder = new DocxBuilder();
	}
	
	
	public static DocxBuilder getInstance(){
		return DocxBuilderHodler.docxBuilder;
	}
	
	
	
	public File build(File tempDocxFile,File targetFile,Object dataObject) {
		//通过freemarker替换内容后都一个文件，用于从新压缩到docx文件中
		File documentXmlFile = new File("document.xml");
		//提取docx中的document.xml文件并输出到doc.xml文件中，因为后面需要用到名称为document.xml的文件
		File tempXmlFile = DocxZipUtil.unZipDocx(tempDocxFile,"word/document.xml", "doc.xml");
		//freemarker通过doc.xml模板文件，生成替换后，带有实际数据到document.xml文件
		try{
			fu.fprint(tempXmlFile, dataObject, documentXmlFile);
		}catch(Exception e){
			logger.error("生成文档失败："+e);
		}
		//将替换好的xml文件放入到docx文件中，并写入到指定位置
		DocxZipUtil.zipDocxFile(documentXmlFile, tempDocxFile, targetFile);
		
		//清理临时文件
		tempXmlFile.delete();
		documentXmlFile.delete();
		
		return targetFile;
	}
	
}
