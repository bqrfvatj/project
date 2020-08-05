package org.ubfs.word.temp.tool;
import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Component;
import org.springframework.util.Assert;
import org.ubfs.word.temp.annoation.CovertEnum;
import org.ubfs.word.temp.annoation.ExeclTitle;

@Component
public class ExeclParseUtil {
	/**
	 * 行位置
	 */
	final static String ROW_KEY = "row";
	/**
	 * 列位置
	 */
	final static String CELL_KEY = "cell";
	
	/**
	 * 列合并数
	 */
	final static String MERGE_ROW_KEY = "merge_row";
	/**
	 * 行合并数
	 */
	final  static String MERGE_CELL_KEY = "merge_cell";
	
	/**
	 * 多级表头分隔符
	 */
	final static  String SPLIT_LABLE = "/";
	
	/**
	 * 数据列开始位置
	 */
	private Integer dataRow = 0;
	
	public Integer getDataRow() {
		return dataRow;
	}
	private void setDataRow(Integer dataRow) {
		this.dataRow = dataRow;
	}
	public  <T> List<T> getDataList(int sheetNum, InputStream in, Class<?> clazz){
		try {
			BufferedInputStream inputStream = this.fromatVaildata(in);
			Sheet sheet = getSheet(sheetNum,inputStream);
			return getDataList(sheet,clazz);
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
	}
	public <T> List<T> getDataList(InputStream in,Class<?> clazz){
		try {
			BufferedInputStream inputStream = this.fromatVaildata(in);
			Sheet sheet = getSheet(0,inputStream);
			return getDataList(sheet,clazz);
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
		
	}
	/**
	 * 格式校验
	 * @author  taolq
	 * @date    2020年5月14日
	 * @time    下午2:16:42
	 * @param in
	 */
	private BufferedInputStream fromatVaildata(InputStream in) {
		BufferedInputStream inputStream = new BufferedInputStream(in);
		boolean isExcel = FileVaildata.isExcel(inputStream);
		Assert.isTrue(isExcel,"请上传Excel格式文件");
		return inputStream;
	}
    
	/**
	 * 获取填充实体
	 * @param sheet
	 * @param clazz
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public  <T> List<T> getDataList(Sheet sheet ,Class<?> clazz) {
		List<T> tList = new ArrayList<T>();
		try {
			int lastRowNum = sheet.getLastRowNum();
			Field[] fields = clazz.getDeclaredFields();
			//判断是否第一次赋值对象
			boolean  firstAdd = true;
			for (Field field : fields) {
				ExeclTitle annotation = field.getAnnotation(ExeclTitle.class);
				if (annotation != null) {
					String fieldName = annotation.value();
					Map<String, Integer> position = getMerFieldPosition(sheet, fieldName);
					if (position != null && position.size() > 0) {
						int dataRow = position.get(ROW_KEY) + position.get(MERGE_ROW_KEY);
						this.setDataRow(dataRow);
						int dataCell = position.get(CELL_KEY);
						//数据列下标
						int row = 0;
						for (int i = dataRow; i <=lastRowNum; i++) {
							Object obj = firstAdd ? clazz.newInstance() : tList.get(row);
							if(sheet.getRow(i)==null) continue;
							Cell cell = sheet.getRow(i).getCell(dataCell);
							field.setAccessible(true);
							field.set(obj, getCellStringVal(cell)+"");
							if(firstAdd) {
								tList.add((T) obj);
							}
							row ++;
						}
					}
				}
				firstAdd = tList.size() > 0 ? false : true;
			}
			return tList;
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
	}

	private  Sheet getSheet(int sheetNum, InputStream in) {
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(in);
			return workbook.getSheetAt(sheetNum);
		} catch (Exception e) {
			try {
				if(workbook != null) {
					workbook.close();
				}
			} catch (IOException e1) {
				throw new RuntimeException(e1.getMessage());
			}
			return getHssfSheet(sheetNum,in);
		}
	}
	
	@SuppressWarnings("resource")
	private  Sheet getHssfSheet(int sheetNum, InputStream in) {
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			return workbook.getSheetAt(sheetNum);
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
	}
    /**
     * 获取单元格合并数
     * @param cell
     * @param sheet
     * @return
     */
	public  int GetMergeCell(Cell cell, Sheet sheet) {
		int mergeSize = 1;
		List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
		for (CellRangeAddress cellRangeAddress : mergedRegions) {
			if (cellRangeAddress.isInRange(cell)) {
				// 获取合并的行数
				mergeSize = cellRangeAddress.getLastColumn() - cellRangeAddress.getFirstColumn() + 1;
				break;
			}
		}
		return mergeSize;
	}
   /**
	* 获取单元格合并数
	* @param cell
	* @param sheet
	* @return
	*/
	public  int GetMergeRow(Cell cell, Sheet sheet) {
	int mergeSize = 1;
	List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
	for (CellRangeAddress cellRangeAddress : mergedRegions) {
		if (cellRangeAddress.isInRange(cell)) {
			// 获取合并的列数
			mergeSize =	cellRangeAddress.getLastRow()-cellRangeAddress.getFirstRow()+1;
			break;
		}
	}
	return mergeSize;
	}
	
	private  Object getCellStringVal(Cell cell) {
		if(cell ==null){
			return StringUtils.EMPTY;
		}
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                return new BigDecimal(cell.getNumericCellValue());
            case STRING:
                return cell.getStringCellValue().trim();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return String.valueOf(cell.getNumericCellValue());
            case BLANK:
                return "";
            case ERROR:
                return String.valueOf(cell.getErrorCellValue());
            default:
                return StringUtils.EMPTY;
        }
    }
    /**
     * 获取字段坐标
     * @param sheet
     * @param fieldName
     * @return
     */
	private  Map<String, Integer> getFieldPosition(Sheet sheet, String fieldName) {
		Map<String, Integer> hashMap = new HashMap<String, Integer>();
		try {
			end: for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
				Row row = sheet.getRow(rowNum);
				short cellNum = row == null ? 0 : row.getLastCellNum();
				for (int i = 0; i < cellNum; i++) {
					Cell cell = row.getCell(i);
					int mergeNum = cell==null ? 0 :GetMergeCell(cell, sheet);
					int mergeRow = cell==null ? 0 :GetMergeRow(cell, sheet);
					if (cell != null && getCellStringVal(cell).equals(fieldName)) {
						hashMap.put(ROW_KEY, rowNum);
						hashMap.put(CELL_KEY, i);
						hashMap.put(MERGE_CELL_KEY, mergeNum);
						hashMap.put(MERGE_ROW_KEY, mergeRow);
						break end;
					}
				}
			}

		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
		return hashMap;
	}


	/**
	 * 指定范围搜索
	 * 
	 * @param sheet
	 * @param firstRowNum
	 * @param firstCellNum
	 * @param LastCellNum
	 * @param fieldName
	 * @return
	 */
	private  Map<String, Integer> getCellByRange(Sheet sheet, int firstRowNum, int firstCellNum, int LastCellNum,
			String fieldName) {
		Map<String, Integer> hashMap = new HashMap<String, Integer>();
		try {
			end: for (int rowNum = firstRowNum; rowNum < sheet.getLastRowNum(); rowNum++) {
				Row row = sheet.getRow(rowNum);
				if (row == null) {
					return hashMap;
				}
				for (int i = firstCellNum; i < firstCellNum + LastCellNum; i++) {
					Cell cell = row.getCell(i);
					int mergeNum = cell==null ? 0 :GetMergeCell(cell, sheet);
					int mergeRow = cell==null ? 0 :GetMergeRow(cell, sheet);
					if (cell != null && getCellStringVal(cell).equals(fieldName)) {
						hashMap.put(ROW_KEY, rowNum);
						hashMap.put(CELL_KEY, i);
						hashMap.put(MERGE_CELL_KEY, mergeNum);
						hashMap.put(MERGE_ROW_KEY, mergeRow);
						break end;
					}
				}
			}
			return hashMap;
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
	}
    /**
     * 获取多级表头坐标
     * @param sheet
     * @param fieldName
     * @return
     */
	private  Map<String, Integer> getMerFieldPosition(Sheet sheet, String fieldName) {
		Map<String, Integer> cellMap = new HashMap<String, Integer>();
		try {
			if (fieldName.indexOf(SPLIT_LABLE) > -1) {
				String[] split = fieldName.split(SPLIT_LABLE);
				int num = 0;
				cellMap = getFieldPosition(sheet,split[num]);
				if(cellMap==null || cellMap.size() ==0) {
					throw new RuntimeException("无法搜索到表头标题为【"+split[num]+"】的数据，请核对!");
				}
				while(num < split.length) {
					if(cellMap.get(MERGE_CELL_KEY) > 1) {
						num ++;
						cellMap = getCellByRange(sheet,cellMap.get(ROW_KEY),cellMap.get(CELL_KEY),cellMap.get(MERGE_CELL_KEY),split[num]);
						if(cellMap==null || cellMap.size() ==0) {
							throw new RuntimeException("无法搜索到表头标题为【"+split[num]+"】的数据，请核对!");
						}
					}else {
						num = split.length;
					}
				}
			}else {
				cellMap = getFieldPosition(sheet,fieldName);
			}

		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
		return cellMap;
	}
	
	
	
	/**
	 * 类型转换
	 * @param sources
	 * @param target
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public  <T> T convertModelEntity(Object sources,Class<T> target){
		try {
			Field[] sfields = sources.getClass().getDeclaredFields();
			Field[] tfields = target.getDeclaredFields();
			Object targetObj = target.newInstance();
			for(Field sfiled : sfields) {
				for(Field tfiled : tfields) {
					ExeclTitle annotation = sfiled.getAnnotation(ExeclTitle.class);
					if(annotation ==null) continue;
					if(tfiled.getName().equals(sfiled.getName())) {
						String typeName = tfiled.getType().getTypeName();
						sfiled.setAccessible(true);
						tfiled.setAccessible(true);
						//数据源
						Object value = sfiled.get(sources);
						if(StringUtils.isEmpty(value.toString())) continue;
						this.setValue(typeName, tfiled, targetObj, value.toString());
					}
				}
			}
			return (T) targetObj;
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
	}
	
	private void setValue(String typeName, Field filed,Object target,String value) {
		try {
			CovertEnum covertEnum = filed.getAnnotation(CovertEnum.class);
			if(covertEnum != null) {
				Object enumValue = this.convertEnumValue(covertEnum, value.toString());
				filed.set(target, enumValue);
			}else {
			 switch(typeName) {
				case  "java.lang.String" : filed.set(target, value);break;
				case  "java.lang.Double" : filed.set(target, convertAmtFormat(value));break;
				case  "java.lang.Integer" : filed.set(target,Double.valueOf(value).intValue());break;
				case  "int" : filed.set(target, Double.valueOf(value).intValue());break;
				case  "double" : filed.set(target, convertAmtFormat(value));break;
				default : filed.set(target, value);break;
			  }
			}
		 
		} catch (Exception e) {
			throw new RuntimeException("数据类型转换错误"+e.getMessage() + "at :"+filed.getName());
		}
	}
	/**
	 * 转换成金额格式
	 * @param value
	 * @return
	 */
	private Double convertAmtFormat(String value) {
		DecimalFormat df = new DecimalFormat("#.00");
		String format = df.format(Double.valueOf(value));
		return Double.valueOf(format);
	}
	
	/**
	 * 转换枚举值
	 * @param annotation
	 * @param name
	 * @return
	 */
	public  Object convertEnumValue(CovertEnum annotation,String name) {
		try {
			Class<?> enumClass = annotation.value();
			Object instance = enumClass.getEnumConstants();
			String methodName = annotation.methodName();
			Method method = enumClass.getDeclaredMethod(methodName, String.class);
			method.setAccessible(true);
			return method.invoke(instance,name);
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
	}

}
