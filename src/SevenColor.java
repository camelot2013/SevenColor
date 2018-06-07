import java.io.File;
import java.util.List;
import java.util.Map;

public class SevenColor {

	public static void main(String[] args) {
		// TODO 自动生成的方法存根
		String excelPath = "D:\\tmp\\区域产品构成维度-品类分解.xlsx";
		File excelFile = new File(excelPath);
		ExcelOperation excelOp1 = new ExcelOperation(excelFile, "xlsx");
		try {
			List<Map<String, Object>> sheetContent = excelOp1.readExcelContent("Sheet1",  2);
			for(int i=0; i<sheetContent.size(); i++){
				System.out.println(sheetContent.get(i).toString());
			}
		} catch (Exception e) {
			// TODO 自动生成的 catch 块
			e.printStackTrace();
		}
	}
	

}
