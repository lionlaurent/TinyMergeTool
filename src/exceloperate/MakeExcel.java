package exceloperate;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MakeExcel extends ComparerVersion{
	public static void main(String[] args) throws Exception {
		
		makeExcel("test1.xls","test2.xls","src/resources/makeexcel_01.xls");
		
	}
	
	
	
	public static void makeExcel(String ver16,String ver17,String path) throws WriteException,IOException{
        
		
		
		//创建工作薄
		WritableWorkbook workbook = null;
		
		try{
			// 创建一个Excel文件对象
            workbook = Workbook.createWorkbook(new File(path));
            // 创建Excel第一个选项卡对象
            WritableSheet sheet = workbook.createSheet("Sheet 1", 0);
            // 设置表头，第一行内容
            // Label参数说明：第一个是列，第二个是行，第三个是要写入的数据值，索引值都是从0开始
            Label label1 = new Label(0, 0, "集团代码");// 对应为第1列第1行的数据
            Label label2 = new Label(1, 0, "航空公司");// 对应为第2列第1行的数据
            Label label3 = new Label(2, 0, "类别");// 对应为第3列第1行的数据
            Label label4 = new Label(3, 0, "航线");// 对应为第4列第1行的数据
            // 添加单元格到选项卡中
            sheet.addCell(label1);
            sheet.addCell(label2);
            sheet.addCell(label3);
            sheet.addCell(label4);
            String[][] mrglst = mergedata(ver17,ver16);
            for(int i=0;i<mrglst.length;i++){
            	for(int j=0;j<mrglst[1].length;j++){
            		sheet.addCell(new Label(i, j + 1, mrglst[i][j]));
            	}
            }
            
            // 遍历集合并添加数据到行，每行对应一个对象
//            for (int i = 0; i < list.size(); i++) {
//                    Customer customer = list.get(i);
//                    // 表头占据第一行，所以下面行数是索引值+1
//                    // 跟上面添加表头一样添加单元格数据，这里为了方便直接使用链式编程
//                    sheet.addCell(new Label(0, i + 1, customer.getName()));
//                    sheet.addCell(new Label(1, i + 1, customer.getAge().toString()));
//                    sheet.addCell(new Label(2, i + 1, customer.getTelephone()));
//                    sheet.addCell(new Label(3, i + 1, customer.getAddress()));
//            }
            // 写入数据到目标文件
            workbook.write();
		}
		catch (Exception e){
			 e.printStackTrace();
		}
        workbook.close();

    }
    
}
