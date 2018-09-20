package com.sbr.docx.util;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.apache.log4j.Logger;


/**
 *
 * @create 2009-7-25 下午06:17:33
 * @since
 */
public class DocxZipUtil {

	
	private static Logger logger = Logger.getLogger(DocxZipUtil.class);
	
	/**
	 * 
	  * 将docx中指定条目解压出来
	  * @param zip docx文件
	  * @param entryName 要解压出来到文件
	  * @param targetFileName 解压到文件写入到哪个新文件中
	  * @return File 解压到新文件
	  * @author 曹传红
	  * @date 2018年9月20日 下午6:16:07
	  * @version 1.0
	 */
	public static File unZipDocx(File zip, String entryName,String targetFileName) {
		BufferedInputStream is = null;
		BufferedOutputStream os = null;
		File entryFile = new File(targetFileName);
		ZipFile zipFile = null;
		try {
			zipFile = new ZipFile(zip.getAbsolutePath());
			ZipEntry entry = zipFile.getEntry(entryName);
			is = new BufferedInputStream(zipFile.getInputStream(entry));
			os = new BufferedOutputStream(new FileOutputStream(entryFile));
			int b = -1;
			while ((b = is.read()) != -1) {
				os.write(b);
			}
		} catch (IOException e) {
			logger.error("获取zip文件失败");
		} finally {
			try {
				if(zipFile!=null){
					zipFile.close();
				}
				if (is != null) {
					is.close();
				}
				if(os!=null){
					os.close();
				}
			} catch (IOException e) {
				logger.error("文件流关闭异常");
			}

		}
		return entryFile;
	}
	
	
	/**
	 * 
	  * <p>方法描述</p>
	  * @param documentFile freemarker替换好的xml文件
	  * @param docxFile  提供到docx文件
	  * @param targetFile 生成到新docx文件
	  * @return void
	  * @author 曹传红
	  * @date 2018年9月20日 下午6:14:06
	  * @version 1.0
	 */
	public static void zipDocxFile(File documentFile, File docxFile,File targetFile){
		ZipFile zipFile = null;
		ZipOutputStream zipout = null;
		try {
			zipFile = new ZipFile(docxFile);
			zipout = new ZipOutputStream(new FileOutputStream(targetFile));
			Enumeration<? extends ZipEntry> zipEntrys = zipFile.entries();
			int len = -1;
			byte[] buffer = new byte[1024];
			while (zipEntrys.hasMoreElements()) {
				ZipEntry next = zipEntrys.nextElement();
				InputStream is = zipFile.getInputStream(next);
				zipout.putNextEntry(new ZipEntry(next.toString()));
				if ("word/document.xml".equals(next.toString())) {
					InputStream in = new FileInputStream(documentFile);
					while ((len = in.read(buffer)) != -1) {
						zipout.write(buffer, 0, len);
					}
					if(in!=null){
						in.close();
					}
				} else {
					while ((len = is.read(buffer)) != -1) {
						zipout.write(buffer, 0, len);
					}
					if(is!=null){
						is.close();
					}
				}
			}
		} catch (FileNotFoundException e) {
			logger.error("文件未找到:"+e);
		} catch (ZipException e) {
			logger.error("文件解压缩异常:"+e);
		} catch (IOException e) {
			logger.error("I/O异常:"+e);
		} finally {
			try {
				if(zipout!=null) {
					//关闭流
					zipout.close();
				}
				if(zipFile!=null) {
					//关闭压缩文件
					zipFile.close();
				}
			} catch (IOException e) {
				logger.error("文件流关闭异常:"+e);
			}
			
		}
	}

}
