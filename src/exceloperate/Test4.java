package exceloperate;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class Test4 extends Test3{
	
	public static void main(String[] args) throws Exception {
		
		String path = "src/resources/lexicon_2.xls";
		Workbook workbook = Workbook.getWorkbook(new File(path)); 
        Sheet sheet = workbook.getSheet(0);
        int col = sheet.getColumns();
        int row = sheet.getRows();
        for(int i = 0;i<col;i++){
        	for (int j = 0;j<row;j++){
        		int e = sheet.getCell(i,j).getContents().indexOf("=");
        		System.out.println(sheet.getCell(i,j).getContents().substring(e + 1).trim());
        	}
        }
        workbook.close();
		
		String[] p4 = new String[4];
		String[] p5 = new String[2];
		String[] p6 = new String[4];
		p4[0] = "src/resources/例子4-滑梯启动--不涉及左右件/5221C53000G20_J_前左应急门滑梯启动机构_20190730.xls";
		p4[1] = "src/resources/例子4-滑梯启动--不涉及左右件/5221C53000G21_G_前左应急门滑梯启动机构_20190730.xls";
		p4[2] = "src/resources/例子4-滑梯启动--不涉及左右件/5221C53000G22_G_前左应急门滑梯连接机构_20190730.xls";
		p4[3] = "src/resources/例子4-滑梯启动--不涉及左右件/5221C53000G23_C_前左应急门滑梯连接机构_20190730.xls";
		p5[0] = "src/resources/例子5-观察窗/5621C04000G20_F_中后机身旅客观察窗组件_20190808.xls";
		p5[1] = "src/resources/例子5-观察窗/5621C04000G22_E_中后机身旅客观察窗组件_20190808.xls";
		p6[0] = "src/resources/例子6-导向槽/5221C24000G20_J_导向槽_20190121.xls";
		p6[1] = "src/resources/例子6-导向槽/5221C24000G21_F_前左应急门导向槽组件_20190121.xls";
		p6[2] = "src/resources/例子6-导向槽/5221C24000G23_D_前左应急门导向槽组件_20190121.xls";
		p6[3] = "src/resources/例子6-导向槽/5221C24000G25_D_前左应急门导向槽组件_20190121.xls";
		String savepath = "src/resources/IPD output.xls";
		
		//MakeOutput(p5, savepath);
		
		//GetExample4();
		//GetExample5();
		//GetExample6();
		
		//CopyTemplate();
		//WriteIntoTemplate(f);
		
		
	}
	
	private static void CopyTemplate() throws IOException {
		String url_src =  "src/resources/例子5-观察窗/输出的IPD--模板.xls"; 
		String url_dest =  "src/resources/例子5-观察窗/template.xls";
		File file_src = (new File(url_src));
		File file_dest = (new File(url_dest));
		FileInputStream input = new FileInputStream(file_src); 
		BufferedInputStream inBuff = new BufferedInputStream(input);
		FileOutputStream output = new FileOutputStream(file_dest);
		BufferedOutputStream outBuff=new BufferedOutputStream(output);
		// 缓冲数组
		byte[] b = new byte[1024];  
	    int len;  
	    while ((len =inBuff.read(b)) != -1) {
	    	outBuff.write(b, 0, len); 
	    }  
	    // 刷新此缓冲的输出流   
	    outBuff.flush();  
	    // 关闭流   
	    inBuff.close();  
	    outBuff.close();  
	    output.close();  
	    input.close();  
		
	}
	
	private static void WriteIntoTemplate(String[][] newsheet, String savepath) throws Exception {
		Workbook wb = Workbook.getWorkbook 
	    (new File("src" + System.getProperty("file.separator") + "resources" + System.getProperty("file.separator") + "template.xls")); 
		WritableWorkbook wwb = Workbook.createWorkbook(new File(savepath), wb);
		WritableSheet ws = wwb.getSheet(0);
		WritableFont font1 = new WritableFont(WritableFont.createFont("Arial"),10, WritableFont.NO_BOLD); // 字体样式
        WritableCellFormat wcf1 = new WritableCellFormat(font1);
        wcf1.setBackground(Colour.YELLOW);
        WritableFont font2 = new WritableFont(WritableFont.createFont("宋体"),10, WritableFont.NO_BOLD); // 字体样式
        WritableCellFormat wcf2 = new WritableCellFormat(font2);
		for(int i=0;i<newsheet.length;i++){
			for(int j=0;j<newsheet[0].length;j++){
				if (newsheet[i][j] != "null") {
					ws.addCell(new Label(i,j+2,newsheet[i][j],wcf2));					
				}	
			}
		}
		for(int i=0;i<newsheet[0].length;i++){
				if (newsheet[7][i].equals("zero") ) {
					ws.addCell(new Label(7,i+2,""));					
				}
				else if (newsheet[6][i].equals("highlight")) {
					ws.addCell(new Label(6,i+2,"",wcf1));
			}
		}
		wwb.write();
		wwb.close();
		wb.close();		
	}
	// put Excel data into the String[][] 
	private static String[][] GetSheet(String path) throws Exception {
		Workbook workbook = Workbook.getWorkbook(new File(path)); 
        Sheet sheet = workbook.getSheet(0);
        int col = sheet.getColumns();
        int row = sheet.getRows()-6;
        String[][] sheet1 = new String[col][row];
        for(int i=0;i<col;i++){
        	for (int j=0;j<row;j++){
        		sheet1[i][j] = sheet.getCell(i,j+2).getContents();
        	}
        }
        workbook.close();
        
        return sheet1;
        
	}
	// print 2D String in console
	public static void PrintMatrix2 (String[][] matrix2) {
		for(int i=0;i<matrix2.length;i++){
			for(int j=0;j<matrix2[0].length;j++){
				System.out.println(matrix2[i][j]);
			}
		}
	}
	// print 1D String in console
	public static void PrintMatrix (String[] matrix) {
			for(int i=0;i<matrix.length;i++){
				System.out.println(matrix[i]);
			}
		}
	// print 2D Array in console
	public static void PrintArray2 (int[][] arr2) {
		for(int i=0;i<arr2.length;i++){
			for(int j=0;j<arr2[0].length;j++){
				System.out.println(arr2[i][j]);
			}
		}
	}
	// print 1D Array in console
	public static void PrintArray (int[] arr){
		for(int i=0;i<arr.length;i++){
			System.out.println(arr[i]);
			}		
	}
	// get number of rows of new version which arn't exist in old version
	private static int[] rows_new (String[] ver_old,String[] ver_new) {
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
	private static int[] rows_plug (String[] ver_old,String[] ver_new) {
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
	private static int[][] MergeRows(String[][] ver_old,int[] plug) {
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
	private static int GetRmdr (double number) {
		double p_int = Math.floor(number) ;
	    double p_remain = number-p_int;
	    int rmdr = (int) Math.round(p_remain*1000);
		return rmdr;
	}
	// get integer from a double 
	private static int GetInt (double number) {
		double p_int = Math.floor(number);
		int integer = (int) p_int;
		return integer;
	}
	// merge 2 sheet into one
	private static String[][] MergeSheet (int[][] rowsnumber,String[][] ver_old,String[][] ver_new,int[] rows_new) {
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
	// get the chosen layer data from whole sheet
	private static String[][] GetLayer (String[][] sheet,int layer) {
		List<Integer> lst = new ArrayList<Integer>();
		switch(layer)
		{
		case 1:		
			for (int i = 0;i<sheet[0].length;i++) {
				if (sheet[8][i].substring(0,3).equals("DCI")){
					lst.add(i);
				}
			}
			break;
		case 2:
			for (int i = 0;i<sheet[0].length;i++) {
				int G = sheet[8][i].indexOf("G");
				if(G == -1){
				}
				else if (sheet[8][i].substring(G,G+2).equals("G2")){
					lst.add(i);
				}
			}
			break;
		case 3:
			for (int i = 0;i<sheet[0].length;i++) {
				int G = sheet[8][i].indexOf("G");
				if(G == -1){
				}
				else if (sheet[8][i].substring(G,G+2).equals("G4")){
					lst.add(i);
				}
			}
			break;
		}
		String[][] lr = new String[sheet.length][lst.size()];
		for (int i = 0;i<sheet.length;i++) {
			for (int j = 0;j<lst.size();j++) {
				lr[i][j] = sheet[i][lst.get(j)];
			}
		}
		
		return lr;
		
	}
	
	private static String[] GetL2Num (String[][] s, String[][] l3) {
		String[] l2num = new String[l3[0].length];
		for (int i = 0;i<l3[0].length;i++) {
			String layernum = l3[1][i];
			int p2 = layernum.indexOf('.', 2);
			String layernum2 = layernum.substring(0, p2);
			//System.out.println(layernum2);
			for (int j = 0;j<s[0].length;j++) {
				if (s[1][j].equals(layernum2)) {
					l2num[i] = s[2][j];
					break;
				}
			}
		}
		
		return l2num;
	}
	
	private static String[] GetL2Num2 (String[][] l2_mrg) {
		String[] l2num = new String[l2_mrg[0].length];
		for (int i = 0;i<l2_mrg[0].length;i++) {
			int de = l2_mrg[0][i].lastIndexOf("_de_");
			l2num[i] =  l2_mrg[0][i].substring(de+4);
		}
		
		return l2num;
	}
	
	private static String[][] Remove_R(String[][] lr) {
		List<Integer> lst = new ArrayList<Integer>();
		for (int i = 0;i<lr[0].length;i++) {
			if (lr[2][i].substring(0,2).equals("R_")){

			}
			else {
				lst.add(i);
			}
		}
		String[][] lr_new = new String[lr.length][lst.size()];
		for (int i = 0;i<lr.length;i++) {
			for (int j = 0;j<lst.size();j++) {
				lr_new[i][j] = lr[i][lst.get(j)];
			}
		}
		
		return lr_new;
	}
	
	private static String[][] RmSame_L2 (String[][] lr_old,String[][] lr_new) {
		List<String> lst = new ArrayList<String> ();
		for (int i = 0;i<lr_old[0].length;i++) {
			if (Integer.parseInt(lr_old[5][i])<10){
				lst.add(lr_old[2][i]+"ver_1"+lr_old[3][i]+0+lr_old[5][i]+"row_old"+i);
			}
			else {
				lst.add(lr_old[2][i]+"ver_1"+lr_old[3][i]+lr_old[5][i]+"row_old"+i);
			}
		}
		for (int i = 0;i<lr_new[0].length;i++) {
			if (Integer.parseInt(lr_new[5][i])<10){
				lst.add(lr_new[2][i]+"ver_2"+lr_new[3][i]+0+lr_new[5][i]+"row_new"+i);
			}
			else {
				lst.add(lr_new[2][i]+"ver_2"+lr_new[3][i]+lr_new[5][i]+"row_new"+i);
			}
		}
		Collections.sort(lst);
//		for (int i = 0;i<lst.size();i++) {
//			System.out.println(lst.get(i));
//		}
		for (int i = 0;i<lst.size()-1;i++) {
			int cont1 = lst.get(i).indexOf("row_");
			int cont2 = lst.get(i+1).indexOf("row_");
			int cont3 = lst.get(i).indexOf("ver_");
			int cont4 = lst.get(i+1).indexOf("ver_");
			if (lst.get(i).substring(0, cont3).equals(lst.get(i+1).substring(0, cont4)) && 
				lst.get(i).substring(cont3+5, cont1).equals(lst.get(i+1).substring(cont4+5, cont2))	) {
				lst.remove(i+1);
			}
		}
		String [][] mrgl2 = new String[lr_old.length][lst.size()];
		for (int i =0;i<mrgl2.length;i++) {
			for (int j =0;j<mrgl2[0].length;j++) {
				int row = Integer.parseInt(lst.get(j).substring(lst.get(j).indexOf("row")+7));
				String ver = lst.get(j).substring(lst.get(j).indexOf("row") + 4, lst.get(j).indexOf("row") + 7);
				if (ver.equals("old")) {
					mrgl2[i][j] = lr_old[i][row];
				}
				else {
					mrgl2[i][j] = lr_new[i][row];
				}
			}
			
		}
		
		return mrgl2;	
	}
	
	private static String[][] RmSame_L3 (String[][] l3_old, String[][] l3_new, String[] n3_old, String[] n3_new) {
		List<String> lst = new ArrayList<String> ();
		for (int i = 0;i<n3_old.length;i++) {
			int g = n3_old[i].indexOf("G4");
			n3_old[i] = n3_old[i].substring(0, g+2);
		}
		for (int i = 0;i<n3_new.length;i++) {
			int g = n3_new[i].indexOf("G4");
			n3_new[i] = n3_new[i].substring(0, g+2);
		}
		for (int i = 0;i<l3_old[0].length;i++) {
			if (Integer.parseInt(l3_old[5][i])<10){
				lst.add(n3_old[i]+l3_old[2][i]+"ver_1"+l3_old[3][i]+0+l3_old[5][i]+"row_old"+i);
			}
			else {
				lst.add(n3_old[i]+l3_old[2][i]+"ver_1"+l3_old[3][i]+l3_old[5][i]+"row_old"+i);
			}
		}
		for (int i = 0;i<l3_new[0].length;i++) {
			if (Integer.parseInt(l3_new[5][i])<10){
				lst.add(n3_new[i]+l3_new[2][i]+"ver_2"+l3_new[3][i]+0+l3_new[5][i]+"row_new"+i);
			}
			else {
				lst.add(n3_new[i]+l3_new[2][i]+"ver_2"+l3_new[3][i]+l3_new[5][i]+"row_new"+i);
			}
		}
		Collections.sort(lst);
		for (int i = 0;i<lst.size()-1;i++) {
			int cont1 = lst.get(i).indexOf("row_");
			int cont2 = lst.get(i+1).indexOf("row_");
			int cont3 = lst.get(i).indexOf("ver_");
			int cont4 = lst.get(i+1).indexOf("ver_");
			if (lst.get(i).substring(0, cont3).equals(lst.get(i+1).substring(0, cont4)) && 
				lst.get(i).substring(cont3+5, cont1).equals(lst.get(i+1).substring(cont4+5, cont2))	) {
				
				lst.remove(i+1);
				
			}
		}
		String [][] mrgl3 = new String[l3_old.length][lst.size()];
		for (int j =0;j<mrgl3[0].length;j++) {
				int row = Integer.parseInt(lst.get(j).substring(lst.get(j).indexOf("row")+7));
				String ver = lst.get(j).substring(lst.get(j).indexOf("row") + 4, lst.get(j).indexOf("row") + 7);
				if (ver.equals("old")) {
					mrgl3[0][j] = l3_old[0][row]+"_de_"+n3_old[row];
				}
				else {
					mrgl3[0][j] = l3_new[0][row]+"_de_"+n3_new[row];
				}
		}
		for (int i =1;i<mrgl3.length;i++) {
			for (int j =0;j<mrgl3[0].length;j++) {
				int row = Integer.parseInt(lst.get(j).substring(lst.get(j).indexOf("row")+7));
				String ver = lst.get(j).substring(lst.get(j).indexOf("row") + 4, lst.get(j).indexOf("row") + 7);
				if (ver.equals("old")) {
					mrgl3[i][j] = l3_old[i][row];
				}
				else {
					mrgl3[i][j] = l3_new[i][row];
				}
			}
		}
			
			return mrgl3;
	}
	
	private static String[][] InsertL3 (String[][] l2, String[][] l3) {
		List<String> lst = new ArrayList<String> ();
		int ll2 = l2[0].length;
		int ll3 = l3[0].length;
		for (int i = 0;i<ll2;i++) {
			lst.add("l2_row"+i);
		}
		for (int i = 0;i<ll3;i++) {
			int de_1 = l3[0][ll3-1-i].indexOf("_de_");
			int de_2 = l3[0][ll3-1-i].lastIndexOf("_de_");
			String up_layer = l3[0][ll3-1-i].substring(de_2+4);
			l3[0][ll3-1-i] = l3[0][ll3-1-i].substring(0, de_1);
			int row_insert = 0;
			for (int j = 0;j<ll2;j++) {
				int g = l2[2][ll2-1-j].indexOf("G4");
				if (up_layer.equals(l2[2][ll2-1-j].substring(0, g+2))) {
					row_insert = ll2-1-j;
					break;
				}
			}
			for (int j = 0;j<lst.size();j++) {
				if (lst.get(j).equals("l2_row"+row_insert)) {
					lst.add(j+1, "l3_row"+(ll3-1-i));
					break;
				}
			}
		}
		String [][] mergel2l3 = new String[l2.length][lst.size()];
		for (int i = 0;i<lst.size();i++) {
			if (lst.get(i).substring(0, 2).equals("l2")) {
				int row = Integer.parseInt(lst.get(i).substring(6));
				for (int j = 0;j<l2.length;j++) { 
					 mergel2l3[j][i] = l2[j][row];
				}
			}
			else {
				int row = Integer.parseInt(lst.get(i).substring(6));
				for (int j = 0;j<l2.length;j++) { 
					 mergel2l3[j][i] = l3[j][row];
				}
			}
		}
		
		return mergel2l3;
		
	}
	
	private static String[][] InsertL1 (String[][] l1, String[][] l2l3) {
		int col = l1.length;
		int row_l1 = l1[0].length;
		int row = l1[0].length+l2l3[0].length;
		String[][] mrgall = new String[col][row];
		for (int i = 0;i<col;i++) {
			for (int j = 0;j<row_l1;j++) {
				mrgall[i][j] = l1[i][j];
			}
			for (int j = row_l1;j<row;j++) {
				mrgall[i][j] = l2l3[i][j-row_l1];
			}
		}
		return mrgall;
	}
	
	private static String[][] MergeL1 (String[][] l1_old, String[][] l1_new) {
		int c = l1_old.length;
		int r1 = l1_old[0].length;
		int r2 = l1_old[0].length + l1_new[0].length;
		String[][] mrgl1 = new String[c][r2];
		for (int i = 0;i<c;i++) {
			for (int j = 0;j<r1;j++) {
				mrgl1[i][j] = l1_old[i][j];
			}
			for (int j = r1;j<r2;j++) {
				mrgl1[i][j] = l1_new[i][j-r1];
			}
		}
		return mrgl1;
	}
	
    private static String[] GetVer (String[][] mergesheet, String[][] orgsheet) throws Exception {
		String[] l_num = GetLayerNum (mergesheet);
		int l1_num = 0;
		for (int i = 0;i<l_num.length;i++) {
			if (l_num[i].equals("1")) {
				l1_num = l1_num + 1;
			}
		}
    	String[] ver = new String[mergesheet[0].length];
    	String orgver = orgsheet[2][0];
    	for (int i = 0;i<2*l1_num;i++) {
    		String item_mrg = mergesheet[2][i]+mergesheet[3][i]+mergesheet[5][i];
    		for (int j = 0;j<orgsheet[0].length;j++) {
    			String item_org = orgsheet[2][j]+orgsheet[3][j]+orgsheet[5][j];
    			if (item_org.equals(item_mrg)) {
    				ver[i] = orgver;
    				break;
    			}
    		}
    	}
    	for (int i = 2*l1_num;i<mergesheet[0].length;i++) {
    		int g1 = mergesheet[8][i].indexOf("G");
    		String item_mrg = mergesheet[2][i]+mergesheet[3][i]+mergesheet[5][i]+mergesheet[8][i].substring(0, g1+2);
    		for (int j = 0;j<orgsheet[0].length;j++) {
    			int g2 = orgsheet[8][j].indexOf("G");
    			String item_org = orgsheet[2][j]+orgsheet[3][j]+orgsheet[5][j]+orgsheet[8][j].substring(0, g2+2);
    			if (item_org.equals(item_mrg)) {
    				ver[i] = orgver;
    				break;
    			}
    		}
    	}
    	return ver;
	}
    
    private static String[] MergeVer (String[] ver1, String[] ver2) {
    	String[] mrgver = new String [ver1.length];
    	for (int i = 0;i<ver1.length;i++) {
    		mrgver[i] = ver1[i] + "|" + ver2[i];
    		int n = mrgver[i].indexOf("null");
    		if (n == -1) {
    			
    		}
    		else {
    			mrgver[i] = mrgver[i].substring(0, n) + mrgver[i].substring(n + 4);
    			int in = mrgver[i].lastIndexOf("|");
    			mrgver[i] = mrgver[i].substring(0, in) + mrgver[i].substring(in + 1);
    		}
    	}
    	
    	return mrgver;
    }
	
    private static String[] GetLayerNum (String[][] all) throws Exception {
		String[] l_num = new String[all[0].length];
		for (int i = 0;i<all[0].length;i++) {
			if (all[8][i].substring(0, 3).equals("DCI")) {
				l_num[i] = "1";
			}
			else if (all[8][i].indexOf("G2") != -1) {
				l_num[i] = "2";
			}
			else if (all[8][i].indexOf("G4") != -1) {
				l_num[i] = "3";
			}
		}
		
		return l_num;
		
		
	}
    
	private static String[][] GetIPD (String[][] all, String[] ver) throws Exception {
		String[] l_num = GetLayerNum (all);
		int l1_num = 0;
		for (int i = 0;i<l_num.length;i++) {
			if (l_num[i].equals("1")) {
				l1_num = l1_num + 1;
			}
		}
		String[][] ipd = new String[16][all[0].length + l1_num];
		
		for (int i = 0;i<l1_num;i++) {
			ipd[2][2*i] = "N";
			ipd[5][2*i] = "1";
			ipd[6][2*i] = "blank";
			ipd[7][2*i] = all[2][i];
			ipd[8][2*i] = all[3][i];
			ipd[9][2*i] = "零件名称";
			ipd[10][2*i] = all[4][i];
			ipd[14][2*i] = ver[i];
			ipd[15][2*i] = "REF";
			ipd[7][2*i + 1] = "zero";
			ipd[9][2*i + 1] = "航材等级代码";
			ipd[10][2*i + 1] = "0";	
		}
		for (int i = 0;i<(l_num.length-l1_num);i++) {
			ipd[5][2*l1_num + i] = l_num[l1_num+i];
			ipd[6][2*l1_num + i] = "blank";
			ipd[7][2*l1_num + i] = all[2][l1_num+i];
			ipd[8][2*l1_num + i] = all[3][l1_num+i];
			ipd[9][2*l1_num + i] = "零件名称";
			ipd[10][2*l1_num + i] = all[4][l1_num+i];
			ipd[14][2*l1_num + i] = ver[l1_num+i];
			ipd[15][2*l1_num + i] = all[5][l1_num+i];
		}

		for (int i = 2*l1_num;i<ipd[0].length;i++) {
			if (ipd[7][i].indexOf("G99") != -1) {
				ipd[6][i] = "highlight";
			}
		}
		
		for (int i = 2*l1_num;i<ipd[0].length;i++) {
			int c = ipd[7][i].indexOf("_");
			if (c != -1) {
				ipd[7][i] = ipd[7][i].substring(0, c) + "/" + ipd[7][i].substring(c+1);
			}
		}
		
		return ipd;
	}
	
	
	
	
	private static void GetExample4() throws Exception {
		String savepath = "src/resources/IPD output.xls";
		String[][] a = GetSheet("src/resources/例子4-滑梯启动--不涉及左右件/5221C53000G20_J_前左应急门滑梯启动机构_20190730.xls");
		String[][] b = GetSheet("src/resources/例子4-滑梯启动--不涉及左右件/5221C53000G21_G_前左应急门滑梯启动机构_20190730.xls");
		String[][] c = GetSheet("src/resources/例子4-滑梯启动--不涉及左右件/5221C53000G22_G_前左应急门滑梯连接机构_20190730.xls");
		String[][] d = GetSheet("src/resources/例子4-滑梯启动--不涉及左右件/5221C53000G23_C_前左应急门滑梯连接机构_20190730.xls");
		String[][] a0 = GetLayer(a,1);
		String[][] b0 = GetLayer(b,1);
		String[][] c0 = GetLayer(c,1);
		String[][] d0 = GetLayer(d,1);
		String[][] a1 = GetLayer(a,2);
		String[][] a2 = GetLayer(b,2);
		String[][] a3 = GetLayer(c,2);
		String[][] a4 = GetLayer(d,2);
		String[][] b1 = GetLayer(a,3);
		String[][] b2 = GetLayer(b,3);
		String[][] b3 = GetLayer(c,3);
		String[][] b4 = GetLayer(d,3);
		a1 = Remove_R(a1);
		a2 = Remove_R(a2);
		a3 = Remove_R(a3);
		a4 = Remove_R(a4);
		b1 = Remove_R(b1);
		b2 = Remove_R(b2);
		b3 = Remove_R(b3);
		b4 = Remove_R(b4);
		String [][] l1 = MergeL1(a0, b0);
		l1 = MergeL1(l1, c0);
		l1 = MergeL1(l1, d0);
		String [][] l2 = RmSame_L2(a1, a2); 
		l2 = RmSame_L2(l2, a3); 
		l2 = RmSame_L2(l2, a4); 
		String[] n1 = GetL2Num(a, b1);
		String[] n2 = GetL2Num(b, b2);
		String[][] l3 = RmSame_L3(b1,b2,n1,n2);
		String[] n3 = GetL2Num(c, b3);
		String[] n3_m = GetL2Num2(l3);
		l3 = RmSame_L3(l3,b3,n3_m,n3);
		String[] n4 = GetL2Num(d, b4);
		String[] n4_m = GetL2Num2(l3);
		l3 = RmSame_L3(l3,b4,n4_m,n4);
		
		String[][] m = InsertL3(l2, l3);
		
		String[][] all = InsertL1(l1, m);
				
		String[] ver1 = GetVer(all, a);
		String[] ver2 = GetVer(all, b);
		String[] ver3 = GetVer(all, c);
		String[] ver4 = GetVer(all, d);
		String[] ver = MergeVer(MergeVer(MergeVer(ver1, ver2), ver3), ver4);
		
		String[][] ipd = GetIPD(all, ver);
		
		WriteIntoTemplate(ipd, savepath); 
	}
	
	private static void GetExample5() throws Exception {
		String savepath = "src/resources/IPD output.xls";
		String[][] a = GetSheet("src/resources/例子5-观察窗/5621C04000G20_F_中后机身旅客观察窗组件_20190808.xls");
		String[][] b = GetSheet("src/resources/例子5-观察窗/5621C04000G22_E_中后机身旅客观察窗组件_20190808.xls");
		String[][] a0 = GetLayer(a, 1);
		String[][] b0 = GetLayer(b, 1);
		String[][] a1 = GetLayer(a,2);
		String[][] a2 = GetLayer(b,2);
		String[][] b1 = GetLayer(a,3);
		String[][] b2 = GetLayer(b,3);
		a1 = Remove_R(a1);
		a2 = Remove_R(a2);
		b1 = Remove_R(b1);
		b2 = Remove_R(b2);
		String [][] l1 = MergeL1(a0, b0);
		String [][] l2 = RmSame_L2(a1, a2);
		String[] n1 = GetL2Num(a, b1);
		String[] n2 = GetL2Num(b, b2);
		String[][] l3 = RmSame_L3(b1,b2,n1,n2);
		String[][] m = InsertL3(l2, l3);
		String[][] all = InsertL1(l1, m);
		String[] ver1 = GetVer(all, a);
		String[] ver2 = GetVer(all, b);
		String[] ver = MergeVer(ver1, ver2);
		String[][] ipd = GetIPD(all, ver);
	    //PrintMatrix2(ipd);
		WriteIntoTemplate(ipd, savepath);
	}
		
	private static void GetExample6() throws Exception {
		String savepath = "src/resources/IPD output.xls";
		String[][] a = GetSheet("src/resources/例子6-导向槽/5221C24000G20_J_导向槽_20190121.xls");
		String[][] b = GetSheet("src/resources/例子6-导向槽/5221C24000G21_F_前左应急门导向槽组件_20190121.xls");
		String[][] c = GetSheet("src/resources/例子6-导向槽/5221C24000G23_D_前左应急门导向槽组件_20190121.xls");
		String[][] d = GetSheet("src/resources/例子6-导向槽/5221C24000G25_D_前左应急门导向槽组件_20190121.xls");
		String[][] a0 = GetLayer(a,1);
		String[][] b0 = GetLayer(b,1);
		String[][] c0 = GetLayer(c,1);
		String[][] d0 = GetLayer(d,1);
		String[][] a1 = GetLayer(a,2);
		String[][] a2 = GetLayer(b,2);
		String[][] a3 = GetLayer(c,2);
		String[][] a4 = GetLayer(d,2);
		String[][] b1 = GetLayer(a,3);
		String[][] b2 = GetLayer(b,3);
		String[][] b3 = GetLayer(c,3);
		String[][] b4 = GetLayer(d,3);
		a1 = Remove_R(a1);
		a2 = Remove_R(a2);
		a3 = Remove_R(a3);
		a4 = Remove_R(a4);
		b1 = Remove_R(b1);
		b2 = Remove_R(b2);
		b3 = Remove_R(b3);
		b4 = Remove_R(b4);
		String [][] l1 = MergeL1(a0, b0);
		l1 = MergeL1(l1, c0);
		l1 = MergeL1(l1, d0);
		String [][] l2 = RmSame_L2(a1, a2); 
		l2 = RmSame_L2(l2, a3); 
		l2 = RmSame_L2(l2, a4); 
		String[] n1 = GetL2Num(a, b1);
		String[] n2 = GetL2Num(b, b2);
		String[][] l3 = RmSame_L3(b1,b2,n1,n2);
		String[] n3 = GetL2Num(c, b3);
		String[] n3_m = GetL2Num2(l3);
		l3 = RmSame_L3(l3,b3,n3_m,n3);
		String[] n4 = GetL2Num(d, b4);
		String[] n4_m = GetL2Num2(l3);
		l3 = RmSame_L3(l3,b4,n4_m,n4);
		
		String[][] m = InsertL3(l2, l3);
		
		String[][] all = InsertL1(l1, m);
				
		String[] ver1 = GetVer(all, a);
		String[] ver2 = GetVer(all, b);
		String[] ver3 = GetVer(all, c);
		String[] ver4 = GetVer(all, d);
		String[] ver = MergeVer(MergeVer(MergeVer(ver1, ver2), ver3), ver4);
		
		String[][] ipd = GetIPD(all, ver);
		
		WriteIntoTemplate(ipd, savepath); 
		
	}
	
	protected static void MakeOutput(String[] path, String savepath) throws Exception {
		int n = path.length;
		List<String[][]> lst = new ArrayList<String[][]> ();
		for (int i = 0;i<n;i++) {
			lst.add(GetSheet(path[i]));
		}
		for (int i = 0;i<n;i++) {
			lst.add(GetLayer(lst.get(i),1));
		}
		for (int i = 0;i<n;i++) {
			lst.add(GetLayer(lst.get(i),2));
		}
		for (int i = 0;i<n;i++) {
			lst.add(GetLayer(lst.get(i),3));
		}
		for (int i = 2*n;i<4*n;i++) {
			lst.set(i, Remove_R(lst.get(i)));
		}
		String[][] l1 = MergeL1(lst.get(n), lst.get(n+1));
		for (int i = n+2;i<2*n;i++) {
			l1 = MergeL1(l1, lst.get(i));
		}
		String [][] l2 = RmSame_L2(lst.get(2*n), lst.get(2*n+1));
		for (int i = 2*n+2;i<3*n;i++) {
			l2 = RmSame_L2(l2, lst.get(i));
		}
		
		String[][] l3 = RmSame_L3
		(lst.get(3*n), lst.get(3*n+1), GetL2Num(lst.get(0), lst.get(3*n)), GetL2Num(lst.get(1), lst.get(3*n+1)));
		for (int i = 3*n+2;i<4*n;i++) {
			l3 = RmSame_L3 (l3, lst.get(i), GetL2Num2(l3), GetL2Num(lst.get(i-3*n), lst.get(i)));
		}
		
		String[][] m = InsertL3(l2, l3);
		
		String[][] all = InsertL1(l1, m);
		
		List<String[]> lstver = new ArrayList<String[]> ();
		for (int i = 0;i<n;i++) {
			lstver.add(GetVer(all, lst.get(i)));
		}
		String[] ver = MergeVer(lstver.get(0), lstver.get(1));
		for (int i = 2;i<n;i++) {
			ver = MergeVer(ver, lstver.get(i));
		}
		
		String[][] ipd = GetIPD(all, ver);
		
		WriteIntoTemplate(ipd, savepath); 
		
	}

	


}
