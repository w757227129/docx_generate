package com.sbr.docx.demo;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import com.sbr.docx.build.DocxBuilder;

public class Demo {
	public static void main(String[] args) {
		//获取docx文档处理实例
		DocxBuilder builder = DocxBuilder.getInstance();
		
		//带freemarker标记模板到docx文件
		String docxFilePath = "docx路径";
		File docxFile = new File(docxFilePath);
		
		//需要生成的文件
		String targetFilePath = "生成到文件路径";
		File targetFile = new File(targetFilePath);
		
		//freemarker替换需要的数据
		Map<String,Object> root = new HashMap<String,Object>();
		//root.put(key, value);
		
		//解压文件docxFile文件，获取到其中到document.xml文件，该文件带有freemarker到标记
		//freemarker根据提供到root数据进行替换
		//替换完成后写入到一个指定到文件targetFile
		builder.build(docxFile, targetFile, root);
		
	}
}
