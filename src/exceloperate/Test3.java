package exceloperate;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import jxl.Sheet;
import jxl.Workbook;

public class Test3 {
	
	public static void main(String args[]) throws Exception{
		
		String[][] a = GetSheet("test1.xls");
		String[][] b = GetSheet("test2.xls");
		int[] c = rows_new(a[3],b[3]) ;
		int[] d = rows_plug(a[3],b[3]) ;
		//PrintArray(d);
		int[][] e = MergeRows(a,d);
		//PrintArray2(e);
		//PrintMatrix2(a);
		String[][] f = MergeSheet(e,a,b,c);
		PrintMatrix2(f);
		
		
	}

	
	
	
	
	// put Excel data into the String[][] 
	private static String[][] GetSheet(String address) throws Exception {
		Workbook workbook = Workbook.getWorkbook(new File("src/resources/"+address)); 
        Sheet sheet = workbook.getSheet(0);
        int col = sheet.getColumns();
        int row = sheet.getRows();
        String[][] sheet1 = new String[col][row];
        for(int i=0;i<col;i++){
        	for (int j=0;j<row;j++){
        		sheet1[i][j] = sheet.getCell(i,j).getContents();
        	}
        }
        workbook.close();
        
        return sheet1;
        
	}
	// print 2D String in console
	private static void PrintMatrix2(String[][] matrix2){
		for(int i=0;i<matrix2.length;i++){
			for(int j=0;j<matrix2[0].length;j++){
				System.out.println(matrix2[i][j]);
			}
		}
	}
	// print 2D Array in console
	private static void PrintArray2(int[][] arr2){
		for(int i=0;i<arr2.length;i++){
			for(int j=0;j<arr2[0].length;j++){
				System.out.println(arr2[i][j]);
			}
		}
	}
	// print 1D Array in console
	private static void PrintArray(int[] arr){
		for(int i=0;i<arr.length;i++){
			System.out.println(arr[i]);
			}		
	}
	// get number of rows of new version which arn't exist in old version
	private static int[] rows_new(String[] ver_old,String[] ver_new){
		List<Integer> whatsnew = new ArrayList<Integer>();
		for(int i=0;i<ver_new.length;i++){
			int y_n = 0;
			for(int j=0;j<ver_old.length;j++){
				if (ver_old[j].equals(ver_new[i])){					
					y_n = 1;
					break;
				}
			}
			if (y_n == 0){
				whatsnew.add(i);
			}
		}
		int s = whatsnew.size();
		int[] rows_new = new int[s];
		for(int i=0;i<s;i++){
			rows_new[i] = whatsnew.get(i);
		}
		
		return rows_new;
	}
	// get number of rows where new rows should be plugged in
	private static int[] rows_plug(String[] ver_old,String[] ver_new){
		List<Integer> wherenew = new ArrayList<Integer>();
		int plugin = 0;
		for(int i=0;i<ver_new.length;i++){
			int y_n = 0;
			for(int j=0;j<ver_old.length;j++){
				if (ver_old[j].equals(ver_new[i])){					
					y_n = 1;
					plugin = j;
					break;
				}
			}
			if(y_n == 0){
				wherenew.add(plugin);
			}
		}
		int s = wherenew.size();
		int[] rows_plug = new int[s];
		for(int i=0;i<s;i++){
			rows_plug[i] = wherenew.get(i);
		}
		
		return rows_plug;
	}
	// get new merged rows number's array
	private static int[][] MergeRows(String[][] ver_old,int[] plug){
		double[] plug_f = new double[plug.length];
		for(int i=0;i<plug.length;i++){
			plug_f[i] = plug[i] + (i+1)*0.001;
			//System.out.println(plug_f[i]);
		}
		double[] rows_old = new double[ver_old[0].length];
		for(int i=0;i<ver_old[0].length;i++){
			rows_old[i] = i;
			//System.out.println(plug_f[i]);
		}
		double[] mrgarr= new double[rows_old.length + plug_f.length];
		for(int i=0;i<plug_f.length;i++){
			mrgarr[i] = plug_f[i];
			//System.out.println(mrgarr[i]);
		}
		for(int i=0;i<rows_old.length;i++){
			mrgarr[plug_f.length+i] = rows_old[i];
			//System.out.println(mrgarr[plug_f.length+i]);
		}
		Arrays.sort(mrgarr);
		int [][] mergerows = new int[mrgarr.length][2];
		for(int i=0;i<mergerows.length;i++){
			if (GetRmdr(mrgarr[i])==0){
				mergerows[i][0] = GetInt(mrgarr[i]);
				mergerows[i][1] = 0;
			}
			else{
				mergerows[i][0] = GetRmdr(mrgarr[i]);
				mergerows[i][1] = 1;
			}
		}
		
		return mergerows;
	}
	// get remainder from a double 
	private static int GetRmdr(double number) {
		double p_int = Math.floor(number) ;
	    double p_remain = number-p_int;
	    int rmdr = (int) Math.round(p_remain*1000);
		return rmdr;
	}
	// get integer from a double 
	private static int GetInt(double number) {
		double p_int = Math.floor(number);
		int integer = (int) p_int;
		return integer;
	}
	// merge 2 sheet into one
	private static String[][] MergeSheet(int[][] rowsnumber,String[][] ver_old,String[][] ver_new,int[] rows_new){
		String[][] newsheet = new String[ver_old.length][rowsnumber.length];
		for(int i=0;i<rowsnumber.length;i++){
			for(int j=0;j<ver_old.length;j++){
				if (rowsnumber[i][1] == 0){
					newsheet[j][i] = ver_old[j][rowsnumber[i][0]];
					//System.out.println(newsheet[j][i]);
				}
				else{
					newsheet[j][i] = ver_new[j][rows_new[rowsnumber[i][0]-1]];
					//System.out.println(newsheet[j][i]);
				}
			}
		}
		
		return newsheet;
	}
		
	
}
