package exceloperate;

import java.io.File;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class ComparerVersion  {
	
	public static void main(String[] args) throws Exception {
		
		
		String[][] a = mergedata("test2.xls","test1.xls");
//		for(int i=0;i<a.length;i++){
//			for(int j=0;j<a[1].length;j++){
//				System.out.println(a[i][j]);
//			}
		int[] b = GetSameItem(a,"test1.xls");
		
	}
	

	public static String[][] mergedata(String ver1,String ver2) throws Exception{
		
        Workbook workbook1 = Workbook.getWorkbook(new File("src/resources/"+ver1)); 
        Sheet sheet1 = workbook1.getSheet(0);
        Workbook workbook2 = Workbook.getWorkbook(new File("src/resources/"+ver2)); 
        Sheet sheet2 = workbook2.getSheet(0);
        
        int [] num17 = new int[100];
        double [] num16 = new double[100];
        double num = 0;
        int k = 0;

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
        	}
        }
       
        for(int i=0;i<k;i++){
        	num16[i] = num16[i] + (i+1)*0.01;
        }
        double [] num16_new = new double[k];
        for(int i=0;i<k;i++){
        	num16_new[i] = num16[i] ;
        }
        
        double [] num_std = new double [sheet1.getRows()-1];
        for(int i=0;i<num_std.length-1;i++){
        	num_std[i] = i+1;
        }
        double [] m = merge(num_std, num16_new);
        String[][] mrglst = new String[4][m.length];
        for(int i=0;i<m.length;i++){
        	if(GetRmdr(m[i])==0){
        		mrglst[0][i]=sheet2.getCell(0,GetInt(m[i])).getContents();
        		mrglst[1][i]=sheet2.getCell(1,GetInt(m[i])).getContents();
        		mrglst[2][i]=sheet2.getCell(2,GetInt(m[i])).getContents();
        		mrglst[3][i]=sheet2.getCell(3,GetInt(m[i])).getContents();
        	}
        	else{
        		mrglst[0][i]=sheet1.getCell(0,num17[GetRmdr(m[i])-1]).getContents();
        		mrglst[1][i]=sheet1.getCell(1,num17[GetRmdr(m[i])-1]).getContents();
        		mrglst[2][i]=sheet1.getCell(2,num17[GetRmdr(m[i])-1]).getContents();
        		mrglst[3][i]=sheet1.getCell(3,num17[GetRmdr(m[i])-1]).getContents();
        	}
        	
        }
        workbook1.close();
        workbook2.close();
        return mrglst;
       
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
	private static int [] GetSameItem(String[][] mergelist,String ver_original) throws Exception{
		
		Workbook workbook = Workbook.getWorkbook(new File("src/resources/"+ver_original)); 
        Sheet sheet = workbook.getSheet(0);
        int NewRows = mergelist[3].length;
        int s = NewRows-(sheet.getRows());
    	int[] size = new int[s];
    	int k = 0;
        for (int i=0; i<sheet.getRows(); i++){
        	String con1 = sheet.getCell(3,i).getContents();
        	int y_n = 0;
        	int nbb = 0;
        	for(int j=0; j<NewRows; j++){
        		String con2 = mergelist[3][j];
        		if (con1.equals(con2)){
        			y_n = 1;
        			nbb = j;
        			break;
        		}        		
        	}
        	if (y_n == 0){
        		size[k] = i;
        		k = k+1; 
        		
        	}
        	else{
        		System.out.println(i);
        	}
        }
        workbook.close();
		return size;
	    }
}
