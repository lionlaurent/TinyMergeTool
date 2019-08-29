package exceloperate;

import java.io.File;

import jxl.Sheet;
import jxl.Workbook;

public class Test2 {
	public static void main(String[] args) throws Exception {
		
		Workbook workbook2 = Workbook.getWorkbook(new File("src/resources/test1.xls")); 
        Sheet sheet2 = workbook2.getSheet(0);
		double a = 2.03;
		System.out.println(GetRmdr(a));
		System.out.println(GetInt(a));
		System.out.println(Math.floor(a));
		System.out.println(a-Math.floor(a));
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
