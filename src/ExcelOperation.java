import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;





import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperation {
	private int columns =0;
	private File excelFile =null;
	String excelType ="";
	ExcelOperation(File file, String type){
		excelFile = file;
		excelType = type;
	}
	
	public List<Map<String, Object>> readExcelContent(String sheetName, int iHeadLines)throws Exception{
		if(excelType.equals("xls")){
			return readExcelContentXls(excelFile,sheetName, iHeadLines);
		}else if(excelType.equals("xlsx")){
			return readExcelContentXlsx(excelFile, sheetName, iHeadLines);
		}
		return null;
	}
	private void putRow2List(Row row, List<Map<String, Object>> sheetContenList)throws Exception{
		if(row ==null)
			return;
		Map<String, Object> rowMap = new HashMap<String, Object>();
		for (short c = 0; c < columns; c++) { // 循环遍历行中的单元格     
			Cell cell = row.getCell((short) c); 
			rowMap.put(""+c, getCellValue(cell));
		}
		sheetContenList.add(rowMap);
	}
//	public String getCellValue(int iRow, int iColumn)throws Exception{
//		String sResult ="";
//		
//		if(iRow<0)
//			return sResult;
//		if(iColumn<0)
//			return sResult;
//		
//		return sResult;
//	}
	private String getCellValue(Cell cell)throws Exception{
		String sResult ="";
		Cell fCell =null;
		if(cell == null)
			return sResult;
		
		if(isMergedRegion(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex())){
			Sheet sheet = cell.getSheet();
			int row = cell.getRowIndex();
			int column = cell.getColumnIndex();
			int sheetMergeCount = sheet.getNumMergedRegions();    
	        
		    for(int i = 0 ; i < sheetMergeCount ; i++){    
		        CellRangeAddress ca = sheet.getMergedRegion(i);    
		        int firstColumn = ca.getFirstColumn();    
		        int lastColumn = ca.getLastColumn();    
		        int firstRow = ca.getFirstRow();    
		        int lastRow = ca.getLastRow();    
		            
		        if(row >= firstRow && row <= lastRow){    
		                
		            if(column >= firstColumn && column <= lastColumn){    
		                Row fRow = sheet.getRow(firstRow);    
		                fCell = fRow.getCell(firstColumn);    
		                break;
		            }    
		        }    
		    }    
		}else{
			fCell = cell;
		}
		switch(fCell.getCellTypeEnum()){
		case STRING:
			sResult = fCell.getStringCellValue();
			break;
		case NUMERIC:
			if(DateUtil.isCellDateFormatted(fCell)){
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd"); 
				sResult = sdf.format(fCell.getDateCellValue()).toString();
			}else{
				sResult = String.valueOf(fCell.getNumericCellValue());
			}
			break;
		case BOOLEAN:
			sResult = String.valueOf(fCell.getBooleanCellValue());
			break;
		default:
			sResult = fCell.getStringCellValue();
			break;
		}
		return sResult;
	}
	private void getColumns(Row row)throws Exception{
		columns = row.getLastCellNum();
	}
	/**  
	* 判断指定的单元格是否是合并单元格  
	* @param sheet   
	* @param row 行下标  
	* @param column 列下标  
	* @return  
	*/  
	private boolean isMergedRegion(Sheet sheet,int row ,int column) {  
		int sheetMergeCount = sheet.getNumMergedRegions();  
		for (int i = 0; i < sheetMergeCount; i++) {  
			CellRangeAddress range = sheet.getMergedRegion(i);  
			int firstColumn = range.getFirstColumn();  
			int lastColumn = range.getLastColumn();  
			int firstRow = range.getFirstRow();  
			int lastRow = range.getLastRow();  
			if(row >= firstRow && row <= lastRow){  
				if(column >= firstColumn && column <= lastColumn){  
					return true;  
				}  
			}  
		}  
		return false;  
	}
	private List<Map<String, Object>> readExcelContentXls(File excelFile, String sheetName, int iHeadLines)throws Exception{
		List<Map<String, Object>> sheetContenList = null;
		// 創建對Excel工作簿文件的引用
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(excelFile));
        // 創建對工作表的引用
       try{
    	   HSSFSheet sheet = workbook.getSheet(sheetName);//讀取第一張工作表 Sheet1
    	   if(sheet ==null)
    		   return null;
           int rows = sheet.getPhysicalNumberOfRows();
           if(rows <2)//如果少于两行(第一行固定为表头)，少于两行实际代表无数据{
        	   return null;
           	
           
           sheetContenList = new ArrayList<Map<String,Object>>();
           for(int iRow =0; iRow <rows; iRow++){
        	   Row row = sheet.getRow(iRow);
        	   if(row !=null){
        		   if(iRow ==iHeadLines){
        			   getColumns(sheet.getRow(iRow));
        		   }
        		   else if(iRow >=iHeadLines)
        			   putRow2List(row, sheetContenList);
        	   }
           }
       }finally{
    	   workbook.close(); 
       }
       return sheetContenList;
	}
	private List<Map<String, Object>> readExcelContentXlsx(File excelFile, String sheetName, int iHeadLines)throws Exception{
		List<Map<String, Object>> sheetContenList = null;
		Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFile));  
		try{
			Sheet sheet = workbook.getSheet(sheetName); // 创建对工作表的引用
			if(sheet ==null)
				return null;
			int rows = sheet.getPhysicalNumberOfRows();
			if(rows <2)//如果少于两行(第一行固定为表头)，少于两行实际代表无数据{
	        	return null;
			sheetContenList = new ArrayList<Map<String,Object>>();
			for(int iRow=0; iRow <rows; iRow++){
				Row row = sheet.getRow(iRow);
	        	if(row !=null){
	        		if(iRow ==iHeadLines-1)
	        			getColumns(sheet.getRow(iRow));
	        		else if(iRow >=iHeadLines)
	        			putRow2List(row, sheetContenList);
	        	}
			}
		}finally{
			workbook.close();
		}
		
		return sheetContenList;
	}
}
