package org.ubfs.word.temp.tool;

import java.io.BufferedInputStream;
import java.io.InputStream;
import java.util.Objects;

import org.apache.poi.poifs.filesystem.FileMagic;

/**
  * 创建时间 ： 2020年5月14日 下午12:39:47
  * 版权所有(C) 2020 深圳雁联数据科技有限公司 
 * 
 * author : taolq
 *    
 */
public class FileVaildata {
	
	/**
	 * 判断是否是excel 文件
	 * @author  taolq
	 * @date    2020年5月14日
	 * @time    下午12:40:54
	 * @param in
	 * @return
	 */
    public static boolean isExcel(InputStream in) {
    	boolean isExcel = false;
    	try {
			FileMagic fileMagic = FileMagic.valueOf(in);
			if(Objects.equals(fileMagic,FileMagic.OLE2) || Objects.equals(fileMagic, FileMagic.OOXML)) {
				isExcel =  true;
			}
		} catch (Exception e) {
			throw new RuntimeException("校验文件格式出现异常");
		}
		return isExcel;
    }
    
    /**
      * 判断文件格式是否是word docx
     * @param in
     * @return
     */
    public static boolean isWordDocx(InputStream in) {
    	boolean isWordDocx = false;
    	try {
    		BufferedInputStream inputStream = new BufferedInputStream(in);
			FileMagic fileMagic = FileMagic.valueOf(inputStream);
			if(Objects.equals(fileMagic,FileMagic.WORD2) || Objects.equals(fileMagic, FileMagic.OOXML)) {
				isWordDocx =  true;
			}
		} catch (Exception e) {
			throw new RuntimeException("校验文件格式出现异常",e);
		}
		return isWordDocx;
    }
}
