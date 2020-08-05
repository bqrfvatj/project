package org.ubfs.word.temp.service;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.springframework.util.Assert;
import org.ubfs.word.temp.annoation.WordParams;
import org.ubfs.word.temp.beans.BaseWordTemp;
import org.ubfs.word.temp.beans.DocInfo;
import org.ubfs.word.temp.beans.ImageInf;
import org.ubfs.word.temp.beans.ParamsBean;
import org.ubfs.word.temp.constant.WordParamsType;
import org.ubfs.word.temp.tool.FileVaildata;

import lombok.extern.slf4j.Slf4j;


/**
 * word 模板处理抽象类
 * 
 * @author taolongqing
 *
 */
@Slf4j
public abstract class AbstractWordTemple {

	/**
	 * 文本标记前缀
	 */
	private final String prefix = ToDBC(setPreFix("{"));
	/**
	 * 文本标记后缀
	 */
	private final String suffix = ToDBC(setSuffix("}"));

	/**
	 * 设置左边模板字符串
	 * 
	 * @param left
	 * @return
	 */
	public String setPreFix(String left) {
		// TODO Auto-generated method stub
		return left;
	}

	/**
	 * 设置右边模板字符串
	 * 
	 * @param right
	 * @return
	 */
	public String setSuffix(String right) {
		// TODO Auto-generated method stub
		return right;
	}
	
	private BufferedInputStream fromatVaildata(InputStream in) {
		BufferedInputStream inputStream = new BufferedInputStream(in);
		boolean isWordDocx = FileVaildata.isWordDocx(inputStream);
		Assert.isTrue(isWordDocx,"请使用docx格式文件作为模板");
		return inputStream;
	}

	/**
	 * 搜索标记符号并替换
	 */
	public final void findLabelAndReplace(BaseWordTemp t) {
		try {
			log.info("开始搜索word 中的文本并准备替换");
			Assert.notNull(t, "模板实体bean不能为空");
			log.info("word 模板文件路径：{}",t.getTempPath());
			t.vaildata();
			fromatVaildata(new FileInputStream(new File(t.getTempPath())));
			//-------------start------------------
			XWPFDocument document = new XWPFDocument(new FileInputStream(new File(t.getTempPath())));
			List<DocInfo> docInfoList = findDocInfByDocument(document, new ArrayList<DocInfo>());
			List<ParamsBean> paramsBeanList = copyPojoFieldToList(t);
			for(DocInfo doc : docInfoList) {
				List<XWPFParagraph> paragraphs = doc.getXWPFParagraphs();
				// 遍历段落
				Iterator<XWPFParagraph> iterator = paragraphs.iterator();
				List<XWPFRun> collectList = new ArrayList<XWPFRun>();
				while (iterator.hasNext()) {
					XWPFParagraph paragraph = iterator.next();
					List<XWPFRun> runs = paragraph.getRuns();
					for (int i = 0; i < runs.size(); i++) {
						String text = "";
					    try {
					    	text = ToDBC(runs.get(i).getText(runs.get(i).getTextPosition()));
						} catch (Exception e) {
							log.warn("标记文本行【{}】不存在或已被删除，无法读取内容",i);
							continue;
						}
						log.info("读取到文本：“{}”",text);
						if (collectList.isEmpty()) {
							// 检测是否包含前缀的文本，开始收集
							if (text != null && text.contains(prefix)) {
								log.info("----------------------------------------------");
								log.info("识别到前缀标记“{}” 开始收集文本",prefix);
								collectList.add(runs.get(i));
							}
						} else {
							// 如果监测已经收集了包含前缀的文本，开始收集后续文本
							collectList.add(runs.get(i));
						}
						// 如果监测到包含后缀的文本且格式正确，加到处理组中，并清空收集器
						if (text != null && text.contains(suffix) && checkTextFormat(text)) {
							log.info("识别到后缀标记“{}” 准备将收集的文本添加到文本组",suffix);
							this.dispatchReplaceTask(document,doc,paramsBeanList, collectList);
							collectList = new ArrayList<XWPFRun>();
						}
					}
				}
			}
			this.beforWriterHandle(document, paramsBeanList);
			this.procWriteNewWord (document, t);
		} catch (Exception e) {
			log.error(" {} 执行异常","findLabelAndReplace",e);
			throw new RuntimeException(e);
		}
	}

	/**
	 * 获取段落从整个文件中
	 * 
	 * @param document
	 * @param arrayList
	 * @return
	 */
	public List<DocInfo> findDocInfByDocument(XWPFDocument document, List<DocInfo> findList) {
		log.info("开始从全文段落中获取段落");
		// 文本段落
		List<XWPFParagraph> paragraphs = document.getParagraphs();
		findList.add(new DocInfo(paragraphs,null));
		log.info("开始从表格中获取段落");
		// 表格段落
		Iterator<XWPFTable> iterator = document.getTablesIterator();
		while (iterator.hasNext()) {
			XWPFTable xwpfTable = iterator.next();
			List<XWPFTableRow> rows = xwpfTable.getRows();
			for (XWPFTableRow row : rows) {
				List<XWPFTableCell> tableCells = row.getTableCells();
				for (XWPFTableCell cell : tableCells) {
					List<XWPFParagraph> table_paragraphs = cell.getParagraphs();
					findList.add(new DocInfo(table_paragraphs,xwpfTable));
				}
			}
		}
		return findList;
	}

	/**
	 * 替换任务调度中心
	 * @param document 
	 * @param doc 
	 * 
	 * @param pojoParamList
	 * @param xWPFRunList
	 */
	private void dispatchReplaceTask(XWPFDocument document, DocInfo doc, List<ParamsBean> pojoParamList, List<XWPFRun> xWPFRunList) {
		try {
			log.info("收集到的包含标记规则的文本组: “{}” 准备进行替换",xWPFRunList);
			String concatText = "";
			for (int i = 0; i < xWPFRunList.size(); i++) {
				XWPFRun run = xWPFRunList.get(i);
				concatText += ToDBC(run.getText(run.getTextPosition()).trim());
				run.setText("", 0);
			}
			// 从bean 取出数据进行替换
			Iterator<ParamsBean> iterator = pojoParamList.iterator();
			while (iterator.hasNext()) {
				ParamsBean bean = iterator.next();
				if (concatText.contains(bean.getField())) {
					String value = bean.getValue().toString();
					log.info("{}匹配到的参数bean类型为“{}”",bean.getField(),bean.getType());
					switch (bean.getType()) {
						case WordParamsType.TEXT:
							concatText = concatText.replace(bean.getField(), value);
							break;
						case WordParamsType.FILE:
							concatText = concatText.replace(bean.getField(), "");
							imageReplaceHander(bean, xWPFRunList);
							break;
						default:
							this.otherReplaceHander(document,bean,concatText,doc,xWPFRunList);
							break;
					}
				}
			}
			int index = getCenterIndex(0, xWPFRunList.size());
			try {
				xWPFRunList.get(index).setText(concatText);
				log.info("标记文本组{}替换成功",xWPFRunList);
			} catch (Exception e) {
				log.warn("标记文本行【{}】不存在或已被删除，无法填充内容",index);
			}
			log.info("--------------------------------------------------");
		} catch (Exception e) {
			throw new RuntimeException("替换文本组发生异常", e);
		}
	}
    

	/**
	 * 转半角的函数(DBC case)<br/>
	 * <br/>
	 * 全角空格为12288，半角空格为32 其他字符半角(33-126)与全角(65281-65374)的对应关系是：均相差65248
	 * 
	 * @param input 任意字符串
	 * @return 半角字符串
	 *
	 */
	protected  String ToDBC(String input) {
		if(input ==null) {
			return "";
		}
		char[] c = input.toCharArray();
		for (int i = 0; i < c.length; i++) {
			if (c[i] == 12288) {
				// 全角空格为12288，半角空格为32
				c[i] = (char) 32;
				continue;
			}
			if (c[i] > 65280 && c[i] < 65375)
				// 其他字符半角(33-126)与全角(65281-65374)的对应关系是：均相差65248
				c[i] = (char) (c[i] - 65248);
		}
		return new String(c);
	}

	/**
	 * 转全角的方法(SBC case)<br/>
	 * <br/>
	 * 全角空格为12288，半角空格为32 其他字符半角(33-126)与全角(65281-65374)的对应关系是：均相差65248
	 * 
	 * @param input 任意字符串
	 * @return 半角字符串
	 *
	 */
	protected  String ToSBC(String input) {
		// 半角转全角：
		if(input ==null) {
			return "";
		}
		char[] c = input.toCharArray();
		for (int i = 0; i < c.length; i++) {
			if (c[i] == 32) {
				c[i] = (char) 12288;
				continue;
			}
			if (c[i] < 127)
				c[i] = (char) (c[i] + 65248);
		}
		return new String(c);
	}

	/**
	 * 图片替换处理
	 * 
	 * @param concatText
	 * @param xWPFRunList
	 */
	private void imageReplaceHander(ParamsBean bean, List<XWPFRun> xWPFRunList) {
		// TODO Auto-generated method stub
		try {
			log.info("开始处理图片替换");
			if(bean.getValue() instanceof ImageInf) {
				ImageInf imageInf = (ImageInf) bean.getValue();
				imageInf.vaildata();
				String imgPath = imageInf.getPath();
				log.info("图片路径：{}",imgPath);
				int index = getCenterIndex(0, xWPFRunList.size());
				String[] split = imgPath.split("\\.", 0);
				FileInputStream is = new FileInputStream(imgPath);
				xWPFRunList.get(index).addPicture(is, getPictureType(split[split.length - 1]), imgPath,
						Units.toEMU(imageInf.getWidth()),
						Units.toEMU(imageInf.getHeight()));
				is.close();
				log.info("图片替换处理成功");
			}else {
				throw new RuntimeException("the bean type is not FILE");
			}
		} catch (Exception e) {
			throw new RuntimeException("图片替换处理出错", e);
		}
	}

	/**
	 * 获取中心位置
	 * 
	 * @return
	 */
	private int getCenterIndex(int count, int total) {
		try {
			int p = (total - count) / 2;
			Double floor = Math.floor(p);
			return total < count ? 0 : floor.intValue();
		} catch (Exception e) {
			throw new RuntimeException("计算中心起始位置出错", e);
		}
	}

	/**
	 * 格式校验
	 * 
	 * @param concatText
	 * @return
	 */
	private boolean checkTextFormat(String concatText) {
		try {
			Pattern r1 = Pattern.compile("(\\" + prefix + "[^\\" + suffix + "]*\\})");
			Pattern r2 = Pattern.compile("\\" + prefix + "");
			Matcher m1 = r1.matcher(concatText);
			Matcher m2 = r2.matcher(concatText);
			int prefixCount = 0;
			int count = 0;
			while (m1.find()) {
				prefixCount++;
			}
			while (m2.find()) {
				count++;
			}
			return count > prefixCount ? false : true;
		} catch (Exception e) {
			throw new RuntimeException("文本正则校验发生错误", e);
		}

	}

	/**
	 * 获取图片类型
	 * 
	 * @param suffix
	 * @return
	 */
	private int getPictureType(String suffix) {
		switch (suffix) {
		case "jpg":
			return XWPFDocument.PICTURE_TYPE_JPEG;
		case "png":
			return XWPFDocument.PICTURE_TYPE_PNG;
		case "gif":
			return XWPFDocument.PICTURE_TYPE_GIF;
		case "bmp":
			return XWPFDocument.PICTURE_TYPE_BMP;
		default:
			return -1;
		}
	}

	/**
	 * 从文件路径中获取文件夹路径
	 * 
	 * @param filePath
	 * @return
	 */
	private String getFolderByFilePath(String filePath) {
		try {
			filePath = filePath.replace("\\", "/");
			String[] split = filePath.split("\\/");
			String folder = filePath.substring(0, filePath.indexOf(split[split.length - 1]));
			return folder;
		} catch (Exception e) {
			throw new RuntimeException("获取文件夹出现错误", e);
		}
	}

	/**
	 * 从对象中提取参数
	 * 
	 * @param entity
	 * @return
	 */
	private List<ParamsBean> copyPojoFieldToList(BaseWordTemp entity) {
		List<ParamsBean> pojoParamList = new ArrayList<ParamsBean>();
		try {
			Field[] fields = entity.getClass().getDeclaredFields();
			for (Field f : fields) {
				WordParams params = f.getAnnotation(WordParams.class);
				if (params != null) {
					f.setAccessible(true);
					ParamsBean joParams = new ParamsBean();
					joParams.setField((prefix + f.getName() + suffix).trim());
					joParams.setType(params.type());
					joParams.setValue(f.get(entity)== null ? "" : f.get(entity));
					joParams.setBaseWordTemp(entity);
					pojoParamList.add(joParams);
				}
			}
		} catch (Exception e) {
			throw new RuntimeException("从对象中提取参数发生异常", e);
		}
		return pojoParamList;
	}

	/**
	 * 导出 word
	 * 
	 * @param document
	 * @param t
	 */
	private void procWriteNewWord (XWPFDocument document, BaseWordTemp base) {
		FileOutputStream fileOut = null;
		try {
			log.info("----------------------------------------------");
			log.info("准备开始导出替换后的word 文件");
			File folder = new File(getFolderByFilePath(base.getOutPath()));
			if (!folder.exists()) {
				log.info("导出文件夹路径不存在，开始自动创建：{}",folder);
				folder.mkdirs();
			}
			log.info("导出文件路径：{}",base.getOutPath());
			fileOut = new FileOutputStream(new File(base.getOutPath()));
			document.write(fileOut);
			fileOut.flush();
			this.afterWriterHandle(document, base);
			log.info("word 文件导出成功");
		} catch (Exception e) {
			log.error("word 文件导出异常",e);
			throw new RuntimeException("文件导出时发生异常", e);
		} finally {
			try {
				fileOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
	/**
      * 其他属性类型替换文本方案
	 * @param document 
    * @param bean
    * @param concatText
	 * @param doc 
	 * @param xWPFRunList 
    */
	public void otherReplaceHander(XWPFDocument document, ParamsBean bean, String concatText, DocInfo doc, List<XWPFRun> xWPFRunList) {
	}

	/**
	 * 导出前处理方法
	 * 
	 * @param pojoParamList
	 * @param document
	 */
	public void beforWriterHandle(XWPFDocument document, List<ParamsBean> pojoParamList) {
		// TODO Auto-generated method stub
	}

	/**
	 * 导出之后处理方法钩子
	 * 
	 * @param document
	 * @param t
	 */
	public void afterWriterHandle(XWPFDocument document, BaseWordTemp t) {
		// TODO Auto-generated method stub
	}

}
