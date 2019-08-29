package exceloperate;

import java.io.File;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class Test {

	public static void main(String[] args) throws Exception {
		
		//1:创建workbook
        Workbook workbook1 = Workbook.getWorkbook(new File("src/resources/test2.xls")); 
        //2:获取第一个工作表sheet
        Sheet sheet1 = workbook1.getSheet(0);
        //3:获取数据
        System.out.println("2017年国内航线数量："+(sheet1.getRows()-1));
        System.out.println("列："+sheet1.getColumns());

        /*     （command+ctrl+/）*/ 

        //1:创建workbook
        Workbook workbook2 = Workbook.getWorkbook(new File("src/resources/test1.xls")); 
        //2:获取第一个工作表sheet
        Sheet sheet2 = workbook2.getSheet(0);
        System.out.println("2016年国内航线数量："+ (sheet2.getRows()-1));
        System.out.println("列："+sheet2.getColumns());
        String [][] merge  = new String[97][4];
        int [] num17 = new int[100];
        double [] num16 = new double[100];
        double [] num_std = new double [49];
        double num = 0;
        int k = 0;
        int kk = 0;
        int kkk = 0;
        for (int i=0; i<sheet1.getRows(); i++){
        	int y_n = 0;
        	
        	Cell cell1 = sheet1.getCell(3,i);
        	String con1 = cell1.getContents();
        	for (int j=0; j<sheet2.getRows(); j++){
        		Cell cell2 = sheet2.getCell(3,j);
        		String con2 = cell2.getContents();
        		if (con1.equals(con2)){
        			num = j;
        			y_n = 1;
        			break;
        		}
        		
        	}
        	if (y_n == 0){
        		num17[k] = i;
        		num16[k] = num;
        		k = k+1;
        		System.out.println("2017年新增的航线有："+con1);
        		System.out.println("所属行数"+num17[k-1]);
        		System.out.println(num16[k-1]);
        	}
        	
        	}
        for(int i=0;i<k;i++){
        	num16[i] = num16[i] + (i+1)*0.01;
        	System.out.println(num16[i]);
        }
        for(int i=0;i<k;i++){
        	System.out.println(num17[i]);
        }
        double [] num16_new = new double[k];
        for(int i=0;i<k;i++){
        	num16_new[i] = num16[i] ;
        }
        for(int i=0;i<49;i++){
        	num_std[i] = i+1;
        	System.out.println(num_std[i]);
        }
        //printArray(merge(num_std, num16_new));
        
        double [] m = merge(num_std, num16_new);
        String [] list = new String[m.length];
        for(int i=0;i<m.length;i++){
        	if(GetRmdr(m[i])==0){
        		list[i]=sheet2.getCell(3,GetInt(m[i])).getContents();
        	}
        	else{
        		list[i]=sheet1.getCell(3,num17[GetRmdr(m[i])-1]).getContents();
        	}
        	System.out.println(list[i]);
        }
        
//        for(int i=0;i<m.length;i++) {
//        	list[i] = GetRmdr(m[i])+"";
//        	System.out.println(list[i]);
//        }
        
        
        
        
        //最后一步：关闭资源
        workbook1.close();
        workbook2.close();
	}
	
	private static double[] merge(double[] a, double [] b) {

        double[] c = new double[a.length+b.length];
        //i用于标记数组a
        int i=0;
        //j用于标记数组b
        int j=0;
        //用于标记数组c
        int k=0;

        //a，b数组都有元素时
        while(i<a.length && j<b.length) {
            if(a[i]<b[j]) {
                c[k++] = a[i++];
            }else {
                c[k++] = b[j++];
            }
        }

        //若a有剩余
        while(i<a.length) {
            c[k++] = a[i++];
        }

        //若b有剩余
        while(j<b.length) {
            c[k++] = b[j++];
        }

        return c;


}
	 private static void printArray(double[]  arr) {
	     int length = arr.length;   
		 for(int i=0;i<length;i++) {
	            System.out.print(arr[i]+ "  ");
	        }
	    }
	 private static int GetRmdr(double number) {
	     double p_int = Math.floor(number) ;
	     double p_remain = number-p_int;
	     int rmdr = (int) Math.round(p_remain*100);
		 return rmdr;
	    }
	 private static int GetInt(double number) {
	     double p_int = Math.floor(number) ;
	     int integer = (int) p_int;
		 return integer;
	    }
	 
}
