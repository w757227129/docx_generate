//获取docx文档处理实例<br />
<strong>DocxBuilder builder = DocxBuilder.getInstance();</strong><br />
<br />
//带freemarker标记模板到docx文件<br />
<strong>String docxFilePath = &quot;docx路径&quot;;<br />
File docxFile = new File(docxFilePath);</strong><br />
<br />
//需要生成的文件<br />
<strong>String targetFilePath = &quot;生成到文件路径&quot;;<br />
File targetFile = new File(targetFilePath);</strong><br />
<br />
//freemarker替换需要的数据<br />
<strong>Map&lt;String,Object&gt; root = new HashMap&lt;String,Object&gt;();<br />
//root.put(key, value);</strong><br />
<br />
//解压文件docxFile文件，获取到其中到document.xml文件，该文件带有freemarker到标记<br />
//freemarker根据提供到root数据进行替换<br />
//替换完成后写入到一个指定到文件targetFile<br />
<strong>builder.build(docxFile, targetFile, root);</strong>
