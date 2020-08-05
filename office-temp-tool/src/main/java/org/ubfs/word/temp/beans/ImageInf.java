package org.ubfs.word.temp.beans;

import java.io.File;

import org.apache.commons.lang.StringUtils;

import lombok.Data;

/**
 * 图片信息bean 
 * @author taolongqing
 *
 */
@Data
public class ImageInf {
	//宽度
	private int width = 200;
	//高度
	private int height = 200;
	//路径
	private String path;
	
	public ImageInf(int width, int height, String path) {
		super();
		this.width = width;
		this.height = height;
		this.path = path;
	}
	public ImageInf() {
		// TODO Auto-generated constructor stub
	}
	public void vaildata() {
		if(StringUtils.isEmpty(path)) {
			throw new RuntimeException("图片地址为空：" + path);
		}
		File file = new File(path);
		if(!file.exists()) {
			throw new RuntimeException("图片文件不存在："+ path);
		}
		if(!file.isFile()) {
			throw new RuntimeException("图片地址异常，必须是文件地址："+ path);
		}
	}
	
}
