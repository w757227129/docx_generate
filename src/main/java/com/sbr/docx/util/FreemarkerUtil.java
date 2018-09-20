package com.sbr.docx.util;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;



public class FreemarkerUtil {

	public Template getTemplate(String name) {
		try {
			// 通过Freemaker的Configuration读取相应的ftl
			Configuration cfg = new Configuration(Configuration.VERSION_2_3_22);
			// 设定去哪里读取相应的ftl模板文件
			// cfg.setDirectoryForTemplateLoading(new File("E:\\freemarker\\"));
			cfg.setClassForTemplateLoading(this.getClass(), "/ftl");
			// 在模板文件目录中找到名称为name的文件
			Template temp = cfg.getTemplate(name);
			return temp;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	public Template getTemplate(String tempFolderPath, String tempName) {
		try {
			// 通过Freemaker的Configuration读取相应的ftl
			Configuration cfg = new Configuration(Configuration.VERSION_2_3_22);
			// 设定去哪里读取相应的ftl模板文件
			cfg.setDirectoryForTemplateLoading(new File(tempFolderPath));
			// 在模板文件目录中找到名称为name的文件
			Template temp = cfg.getTemplate(tempName);
			return temp;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	public Template getTemplate(File TemplateFile) {
		String filePath = TemplateFile.getAbsolutePath();
		String directoryForTemplateLoadingPath = filePath.replace(TemplateFile.getName(), "");
		try {
			Configuration cfg = new Configuration(Configuration.VERSION_2_3_22);
			cfg.setDirectoryForTemplateLoading(new File(directoryForTemplateLoadingPath));
			Template temp = cfg.getTemplate(TemplateFile.getName());
			return temp;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	
	
	public void fprint(File tempFile, Object root,String targetFile) {
		FileWriter out = null;
		try {
			out = new FileWriter(new File(targetFile));
			Template temp = this.getTemplate(tempFile);
			temp.process(root, out);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (TemplateException e) {
			e.printStackTrace();
		} finally {
			try {
				if (out != null)
					out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void fprint(File tempFile, Object root,File targetFile) {
		FileWriter out = null;
		try {
			out = new FileWriter(targetFile);
			Template temp = this.getTemplate(tempFile);
			temp.process(root, out);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (TemplateException e) {
			e.printStackTrace();
		} finally {
			try {
				if (out != null)
					out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
}