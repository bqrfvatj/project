package org.ubfs.word.temp.service.imple;

import java.lang.reflect.Field;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;
import org.ubfs.word.temp.annoation.WordTableParams;
import org.ubfs.word.temp.beans.DocInfo;
import org.ubfs.word.temp.beans.ParamsBean;
import org.ubfs.word.temp.constant.WordParamsType;
import org.ubfs.word.temp.service.AbstractWordTemple;

import lombok.extern.slf4j.Slf4j;

/**
 * 表格循环标签替换处理类
 * 
 * @author taolongqing
 *
 */
@Slf4j
public class TableLoopReplaceHandle extends AbstractWordTemple {
	
	/**
	 * 表格标记前缀
	 */
	private final String table_prefix = ToDBC(setTablePreFix("<"));
	/**
	 * 表格标记后缀
	 */
	private final String table_suffix = ToDBC(setTableSuffix(">"));
	/**
	 * 设置左标记
	 * @param left
	 * @return
	 */
	public String setTableSuffix(String left) {
		// TODO Auto-generated method stub
		return left;
	}
    /**
     * 设置右标记
     * @param right
     * @return
     */
	public String setTablePreFix(String right) {
		// TODO Auto-generated method stub
		return right;
	}
	
	@Override
	public void otherReplaceHander(XWPFDocument document,ParamsBean bean, String concatText,DocInfo doc,List<XWPFRun> xWPFRunList) {
		// TODO Auto-generated method stub
		if(bean.getType()==WordParamsType.LIST && doc.getXwpfTable() != null && concatText.equals(bean.getField())) {
			xWPFRunList.get(0).setText(concatText);
			this.loopLableHandle(doc.getXwpfTable(), bean);
			log.info(" 循环标签表格数据替换完成");
		}
	}
    /**
      * 循环标签处理流程
     * @param xwpfTable
     * @param bean
     */
	public void loopLableHandle(XWPFTable xwpfTable, ParamsBean bean) {
		// TODO Auto-generated method stub
		try {
			log.info(" 程序进入循环标签的处理流程");
			List<Integer> labelIndexList = new ArrayList<Integer>();
			CTTblLayoutType type = xwpfTable.getCTTbl().getTblPr().addNewTblLayout();
			type.setType(STTblLayoutType.AUTOFIT);
			List<XWPFTableRow> rows = xwpfTable.getRows();
			//遍历表格所有列
			for (int i=0;i< rows.size();i++) {
				List<XWPFTableCell> cells = rows.get(i).getTableCells();
				//遍历表格列所有单元格
				for (int y= 0;y<cells.size(); y++) {
					//根据表格数据列表的行数创建 row
					this.createRowsByList(bean,xwpfTable,rows,i,y,labelIndexList);
				}
			}
			//删除标记row
			for(Integer pos : labelIndexList) {
				xwpfTable.removeRow(pos);
				log.info(" 删除循环标签行下标【{}】成功",pos);
			}
		} catch (Exception e) {
			throw new RuntimeException("处理循环标签时异常",e);
		}
	}
	
	

	/**
	 * 根据表格数据列表的行数创建 row
	 * @param pojoParamList
	 * @param xwpfTable 
	 * @param rows
	 * @param cells
	 * @param y 单元格下标
	 * @param i 表格列下标
	 */
    @SuppressWarnings("unchecked")
	private void createRowsByList(ParamsBean bean, XWPFTable xwpfTable, List<XWPFTableRow> rows, int i, int y,List<Integer> labelIndexList) {
			//匹配标记规则的单元格 && 字段类型为list
    		String text = rows.get(i).getTableCells().get(y).getText().trim();
			if(null != text && text.equals(bean.getField()) && bean.getType() == WordParamsType.LIST) {
				XWPFTableRow labelRow = rows.get(i+2); //表格参数标签列 <label>
				log.info("表格第{}行首列内容:{}",i+1, rows.get(i+1).getCell(0).getText());
				log.info("表格第{}行首列内容:{}",i+2, rows.get(i+2).getCell(0).getText());
				int cellSize = labelRow.getTableCells().size();
				//从匹配到的单元格下3行开始循环处理列表数据 ，前两行设置成了规定的表头和参数标记
				int dataIndex = i + 3;
				Object value = bean.getValue();
				 //判断从表格实体中提取出来的value 是否是List 对象的示例
				 if(null !=value && value instanceof List) {
					List<Object> list = (List<Object>) bean.getValue();
					log.info("表格第{}行开始,共计{}行，将列表数据填充到单元格中",dataIndex,list.size());
					final int total = dataIndex + list.size(); //模拟循环列表的长度
					for(int x = dataIndex; x < total; x++) {
						XWPFTableRow createRow = insertNewTableRow(xwpfTable,x,cellSize);
						this.scannTableCellMatchName(cellSize,labelRow,createRow,xwpfTable,list.get(x - dataIndex));
					}
					//删除表格参数标签列
					xwpfTable.removeRow(xwpfTable.getRows().indexOf(labelRow));
					log.info("删除表格参数行下标【{}】单元格成功",i+2);
					//按先进后出的顺序的添加标签row 的下标 
					labelIndexList.add(0,i);
				 }else {
					 log.warn("{} not instanceof list ", value);
					 return;
				 }
				 
			}
		
	}

	/**
      * 匹配列表对象里的标记符号并处理
     * @param cellsSize 数据行长度
     * @param xwpfTableRow 标记了表格属性的Row
     * @param createRow  新增row
     * @param xwpfTable  被编辑的table
     * @param object     参数bean
     */
	private void scannTableCellMatchName(int cellsSize, XWPFTableRow xwpfTableRow,
			XWPFTableRow createRow, XWPFTable xwpfTable, Object object) {
		//填充一行的内容
		for(int f = 0;f< cellsSize;f++) {
			String cellText = xwpfTableRow.getCell(f).getText();
			Iterator<Entry<String, Object>> iterator = copyPojoFieldToMap(object).entrySet().iterator();
			while(iterator.hasNext()) {
				Entry<String, Object> next = iterator.next();
				String value = next.getValue()== null ? "" : next.getValue().toString();
				if(cellText != null && cellText.equals(next.getKey())) {
					createRow = createRow == null ? xwpfTable.createRow() : createRow; 
					this.setTableCell(createRow.getCell(f),value);
			     }
		    }
		}
	}
	/**
	 * 从bean 提取参数到map
	 * @param entity
	 * @return
	 */
	private Map<String,Object> copyPojoFieldToMap(Object entity) {
		Map<String,Object> beanParamsMap = new HashMap<String,Object>();
		try {
			Field[] fields = entity.getClass().getDeclaredFields();
			for (Field f : fields) {
				WordTableParams params = f.getAnnotation(WordTableParams.class);
				if (params != null) {
					f.setAccessible(true);
					beanParamsMap.put((table_prefix + f.getName() + table_suffix).trim(),f.get(entity)== null ? "" : f.get(entity));
				}
			}
			return beanParamsMap;
		} catch (Exception e) {
			throw new RuntimeException("从对象中提取参数发生异常", e);
		}
		
	}

	/**
	 * 插入指定行到表格
	 * @param xwpfTable
	 * @param pos
	 * @param cellSize
	 * @return
	 */
	private XWPFTableRow insertNewTableRow(XWPFTable xwpfTable ,int pos,int cellSize) {
		XWPFTableRow newTableRow = xwpfTable.insertNewTableRow(pos);
		if(newTableRow != null) {
			for(int i=0;i<cellSize;i++) {
				XWPFTableCell createCell = newTableRow.createCell();
				this.setCellWidthAndVAlign(createCell,"6500",STVerticalJc.CENTER,STJc.CENTER);
			}
		}
		newTableRow.setHeight(500);
		return newTableRow;
	}
	
	
	/** 
	* @Description: 设置列宽和垂直对齐方式 
	*/  
	private void setCellWidthAndVAlign(XWPFTableCell cell, String width, STVerticalJc.Enum typeEnum, STJc.Enum vAlign) {
		CTTc cttc = cell.getCTTc();
		CTTcPr cellPr = cttc.isSetTcPr() ? cttc.getTcPr() : cttc.addNewTcPr();
		cellPr.addNewVAlign().setVal(typeEnum);
		cttc.getPList().get(0).addNewPPr().addNewJc().setVal(vAlign);
		CTTblWidth tblWidth = cellPr.isSetTcW() ? cellPr.getTcW() : cellPr.addNewTcW();
		if(!StringUtils.isEmpty(width)){
			tblWidth.setW(new BigInteger(width));
			tblWidth.setType(STTblWidth.DXA);
		}
	}

	/**
	 * 填充单元格内容
	 * @param cell
	 * @param value
	 */
	private void setTableCell(XWPFTableCell cell,Object value) {
		List<XWPFParagraph> paragraphs = cell.getParagraphs();
		for(XWPFParagraph paragraph : paragraphs) {
			Iterator<XWPFRun> iterator = paragraph.getRuns().iterator();
			while(iterator.hasNext()) {
				XWPFRun next = iterator.next();
				next.setText("",0);
			}
		}
		cell.setText(value.toString());
	}

	

}
