import java.io.File;
import java.util.List;
import java.util.Map;

public class SevenColor {

	public static void main(String[] args) {
		// TODO �Զ����ɵķ������
		String excelPath = "D:\\tmp\\�����Ʒ����ά��-Ʒ��ֽ�.xlsx";
		File excelFile = new File(excelPath);
		ExcelOperation excelOp1 = new ExcelOperation(excelFile, "xlsx");
		try {
			List<Map<String, Object>> sheetContent1 = excelOp1.readExcelContent("Sheet1",  2);
			for(int i=0; i<sheetContent1.size(); i++){
				System.out.println(sheetContent1.get(i).toString());
			}
			List<Map<String, Object>> sheetContent2 = excelOp1.readExcelContent("Sheet2");
			for(int i=0; i<sheetContent2.size(); i++){
				System.out.println(sheetContent2.get(i).toString());
			}
		} catch (Exception e) {
			// TODO �Զ����ɵ� catch ��
			e.printStackTrace();
		}
	}
	

}
