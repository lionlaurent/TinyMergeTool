package exceloperate;


import java.io.File;
import java.util.ArrayList;
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

public class Test5 {
	
	public static void main (String[] args) throws Exception {
		
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
		
		makeOutput(p5, savepath);
		
		//GetExample4();
		//GetExample5();
		//GetExample6();
		
		//CopyTemplate();
		//WriteIntoTemplate(f);
		
		
	}
	
	
	
	// put all Excel data into the String[][] 
	private static String[][] getSheet (String path) throws Exception {
		Workbook workbook = Workbook.getWorkbook(new File(path)); 
        Sheet sheet = workbook.getSheet(0);
        int col = sheet.getColumns();
        int row = sheet.getRows()-6;
        String[][] sheet1 = new String[col][row];
        for(int i = 0; i < col; i++) {
        	for (int j = 0; j < row; j++) {
        		sheet1[i][j] = sheet.getCell(i, j+2).getContents();
        	}
        }
        workbook.close();
        
        return sheet1;     
	}
	
	// get the chosen layer data from whole sheet
	private static String[][] getLayer (String[][] sheet, int layer) {
		List<Integer> lst = new ArrayList<Integer>();
		switch(layer)
		{
		case 1:		
			for (int i = 0; i < sheet[0].length; i++) {
				if (sheet[8][i].substring(0, 3).equals("DCI")) {
					lst.add(i);
				}
			}
			break;
		case 2:
			for (int i = 0;i < sheet[0].length; i++) {
				int G = sheet[8][i].indexOf("G");
				if (G != -1) {
					if (sheet[8][i].substring(G, G+2).equals("G2")) {
						lst.add(i);
					}	
				}
			}
			break;
		case 3:
			for (int i = 0; i<sheet[0].length; i++) {
				int G = sheet[8][i].indexOf("G");
				if (G != -1) {
					if (sheet[8][i].substring(G, G+2).equals("G4")) {
						lst.add(i);
					}
				}
			}
			break;
		}
		String[][] lr = new String[sheet.length][lst.size()];
		for (int i = 0; i<sheet.length; i++) {
			for (int j = 0; j<lst.size(); j++) {
				lr[i][j] = sheet[i][lst.get(j)];
			}
		}
		
		return lr;
	}
	
	// remove all the rows which have partnumber begin with "R_" 
	private static String[][] remove_R (String[][] lr) {
		List<Integer> lst = new ArrayList<Integer>();
		for (int i = 0; i < lr[0].length; i++) {
			if (lr[2][i].substring(0,2).equals("R_")) {
			}
			else {
				lst.add(i);
			}
		}
		String[][] lr_new = new String[lr.length][lst.size()];
		for (int i = 0; i < lr.length; i++) {
			for (int j = 0; j < lst.size(); j++) {
				lr_new[i][j] = lr[i][lst.get(j)];
			}
		}
		
		return lr_new;
	}
	
	// merge layer1 into one
	private static String[][] mergeL1 (String[][] l1_old, String[][] l1_new) {
		int c = l1_old.length;
		int r1 = l1_old[0].length;
		int r2 = l1_old[0].length + l1_new[0].length;
		String[][] mrgl1 = new String[c][r2];
		for (int i = 0; i < c; i++) {
			for (int j = 0; j < r1; j++) {
				mrgl1[i][j] = l1_old[i][j];
			}
			for (int j = r1; j < r2; j++) {
				mrgl1[i][j] = l1_new[i][j-r1];
			}
		}
		
		return mrgl1;
	}
	
	// merge layer2 into one & remove those same rows
	private static String[][] mergeL2 (String[][] l2_old, String[][] l2_new) {
		List<String> lst = new ArrayList<String> ();
		for (int i = 0; i < l2_old[0].length; i++) {
			if (Integer.parseInt(l2_old[5][i]) < 10) {
				lst.add(l2_old[2][i] + "ver_1" + l2_old[3][i] + 0 + l2_old[5][i] + "row_old" + i);
			}
			else {
				lst.add(l2_old[2][i] + "ver_1" + l2_old[3][i] + l2_old[5][i] + "row_old" + i);
			}
		}
		for (int i = 0; i < l2_new[0].length; i++) {
			if (Integer.parseInt(l2_new[5][i]) < 10) {
				lst.add(l2_new[2][i] + "ver_2" + l2_new[3][i] + 0 + l2_new[5][i] + "row_new" + i);
			}
			else {
				lst.add(l2_new[2][i] + "ver_2" + l2_new[3][i] + l2_new[5][i] + "row_new" + i);
			}
		}
		Collections.sort(lst);
//		for (int i = 0;i<lst.size();i++) {
//			System.out.println(lst.get(i));
//		}
		for (int i = 0; i < lst.size() - 1; i++) {
			int r1 = lst.get(i).indexOf("row_");
			int r2 = lst.get(i + 1).indexOf("row_");
			int v1 = lst.get(i).indexOf("ver_");
			int v2 = lst.get(i + 1).indexOf("ver_");
			if (lst.get(i).substring(0, v1).equals(lst.get(i + 1).substring(0, v2)) && 
				lst.get(i).substring(v1 + 5, r1).equals(lst.get(i + 1).substring(v2 + 5, r2))) {
				lst.remove(i + 1);
			}
		}
		String [][] mrgl2 = new String[l2_old.length][lst.size()];
		for (int i = 0; i<mrgl2.length; i++) {
			for (int j = 0; j<mrgl2[0].length; j++) {
				int row = Integer.parseInt(lst.get(j).substring(lst.get(j).indexOf("row") + 7));
				String ver = lst.get(j).substring(lst.get(j).indexOf("row") + 4, lst.get(j).indexOf("row") + 7);
				if (ver.equals("old")) {
					mrgl2[i][j] = l2_old[i][row];
				}
				else {
					mrgl2[i][j] = l2_new[i][row];
				}
			}
		}
		
		return mrgl2;	
	}
	
	// get the partnumber in layer2 for layer3's partnumber (when those l2 p_numbr come from original sheet) 
	private static String[] getL2Num (String[][] s, String[][] l3) {
		String[] prt_num = new String[l3[0].length];
		for (int i = 0; i < l3[0].length; i++) {
			int pt2 = l3[1][i].indexOf('.', 2);
			String lr_num = l3[1][i].substring(0, pt2);
			//System.out.println(lyr_num);
			for (int j = 0; j < s[0].length; j++) {
				if (s[1][j].equals(lr_num)) {
					prt_num[i] = s[2][j];
					break;
				}
			}
		}
		
		return prt_num;
	}

	// merge layer3 into one & remove those same rows
	private static String[][] mergeL3 (String[][] l3_old, String[][] l3_new, String[] prtnum_old, String[] prtnum_new) {
		List<String> lst = new ArrayList<String> ();
		for (int i = 0; i < prtnum_old.length; i++) {
			int g = prtnum_old[i].indexOf("G4");
			prtnum_old[i] = prtnum_old[i].substring(0, g + 2);
		}
		for (int i = 0; i < prtnum_new.length; i++) {
			int g = prtnum_new[i].indexOf("G4");
			prtnum_new[i] = prtnum_new[i].substring(0, g + 2);
		}
		for (int i = 0; i < l3_old[0].length; i++) {
			if (Integer.parseInt(l3_old[5][i]) < 10) {
				lst.add(prtnum_old[i] + l3_old[2][i] + "ver_1" + l3_old[3][i] + 0 + l3_old[5][i] + "row_old" + i);
			}
			else {
				lst.add(prtnum_old[i] + l3_old[2][i] + "ver_1" + l3_old[3][i] + l3_old[5][i] + "row_old" + i);
			}
		}
		for (int i = 0; i < l3_new[0].length; i++) {
			if (Integer.parseInt(l3_new[5][i]) < 10) {
				lst.add(prtnum_new[i] + l3_new[2][i] + "ver_2" + l3_new[3][i] + 0 + l3_new[5][i] + "row_new" + i);
			}
			else {
				lst.add(prtnum_new[i] + l3_new[2][i] + "ver_2" + l3_new[3][i] + l3_new[5][i] + "row_new" + i);
			}
		}
		Collections.sort(lst);
		for (int i = 0; i < lst.size() - 1; i++) {
			int r1 = lst.get(i).indexOf("row_");
			int r2 = lst.get(i + 1).indexOf("row_");
			int v1 = lst.get(i).indexOf("ver_");
			int v2 = lst.get(i + 1).indexOf("ver_");
			if (lst.get(i).substring(0, v1).equals(lst.get(i + 1).substring(0, v2)) && 
				lst.get(i).substring(v1 + 5, r1).equals(lst.get(i + 1).substring(v2 + 5, r2))) {
				lst.remove(i + 1);			
			}
		}
		String[][] mrgl3 = new String[l3_old.length][lst.size()];
		for (int i = 0; i < lst.size(); i++) {
			int row = Integer.parseInt(lst.get(i).substring(lst.get(i).indexOf("row") + 7));
			String ver = lst.get(i).substring(lst.get(i).indexOf("row") + 4, lst.get(i).indexOf("row") + 7);
			if (ver.equals("old")) {
				mrgl3[0][i] = l3_old[0][row] + "_de_" + prtnum_old[row];
			}
			else {
				mrgl3[0][i] = l3_new[0][row] + "_de_" + prtnum_new[row];
			}
		}
		for (int i = 1; i < l3_old.length; i++) {
			for (int j = 0; j < lst.size(); j++) {
				int row = Integer.parseInt(lst.get(j).substring(lst.get(j).indexOf("row") + 7));
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
	
	// get the partnumber in layer2 for layer3's partnumber (when those l2 p_numbr come from merged sheet)
	private static String[] getL2Num2 (String[][] l3_mrg) {
		String[] prt_num = new String[l3_mrg[0].length];
		for (int i = 0; i < l3_mrg[0].length; i++) {
			int de = l3_mrg[0][i].lastIndexOf("_de_");
			prt_num[i] = l3_mrg[0][i].substring(de + 4);
		}
		
		return prt_num;
	}
	
	// merge merged layer2 & merged layer3 
	private static String[][] insertL2L3 (String[][] l2, String[][] l3) {
		List<String> lst = new ArrayList<String> ();
		int ll2 = l2[0].length;
		int ll3 = l3[0].length;
		for (int i = 0; i < ll2; i++) {
			lst.add("l2_row" + i);
		}
		for (int i = 0; i < ll3; i++) {
			int de_1 = l3[0][ll3 - 1 - i].indexOf("_de_");
			int de_2 = l3[0][ll3 - 1 - i].lastIndexOf("_de_");
			String prt_num = l3[0][ll3 - 1 - i].substring(de_2 + 4);
			l3[0][ll3 - 1 - i] = l3[0][ll3 - 1 - i].substring(0, de_1);
			int row_insert = 0;
			for (int j = 0; j < ll2; j++) {
				int g = l2[2][ll2 - 1 - j].indexOf("G4");
				if (prt_num.equals(l2[2][ll2 - 1 - j].substring(0, g + 2))) {
					row_insert = ll2 - 1 - j;
					break;
				}
			}
			for (int j = 0;j < lst.size(); j++) {
				if (lst.get(j).equals("l2_row" + row_insert)) {
					lst.add(j + 1, "l3_row" + (ll3 - 1 - i));
					break;
				}
			}
		}
		String[][] mrgl2l3 = new String[l2.length][lst.size()];
		for (int i = 0; i < lst.size(); i++) {
			if (lst.get(i).substring(0, 2).equals("l2")) {
				int row = Integer.parseInt(lst.get(i).substring(6));
				for (int j = 0;j < l2.length; j++) { 
					 mrgl2l3[j][i] = l2[j][row];
				}
			}
			else {
				int row = Integer.parseInt(lst.get(i).substring(6));
				for (int j = 0; j < l2.length; j++) { 
					 mrgl2l3[j][i] = l3[j][row];
				}
			}
		}
		
		return mrgl2l3;
	}
	
	// merge merged layer1 & merged layer2&3
	private static String[][] insertL1 (String[][] l1, String[][] l2l3) {
		int col = l1.length;
		int row = l1[0].length + l2l3[0].length;
		String[][] mrgall = new String[col][row];
		for (int i = 0;  i< col; i++) {
			for (int j = 0; j < l1[0].length; j++) {
				mrgall[i][j] = l1[i][j];
			}
			for (int j = l1[0].length; j < row; j++) {
				mrgall[i][j] = l2l3[i][j - l1[0].length];
			}
		}
		
		return mrgall;
	}
	
	// get layer number of each rows of merged all layers
    private static int HowManyL1 (String[][] all) {
		int l1_length = 0;
		for (int i = 0; i < all[0].length; i++) {
			if (all[8][i].substring(0, 3).equals("DCI")) {
				l1_length = l1_length + 1;
			}
		}
		
		return l1_length;	
	}
	
	//show whether a original sheet version exist in the merged sheet
    private static String[] getVer (String[][] mrgsheet, String[][] orgsheet) throws Exception {
		int l1_length = HowManyL1 (mrgsheet);
    	String[] ver = new String[mrgsheet[0].length];
    	String orgver = orgsheet[2][0];
    	for (int i = 0; i < l1_length; i++) {
    		String item_mrg = mrgsheet[2][i] + mrgsheet[3][i] + mrgsheet[5][i];
    		for (int j = 0; j < orgsheet[0].length; j++) {
    			String item_org = orgsheet[2][j] + orgsheet[3][j] + orgsheet[5][j];
    			if (item_org.equals(item_mrg)) {
    				ver[i] = orgver;
    				break;
    			}
    		}
    	}
    	for (int i = l1_length; i < mrgsheet[0].length; i++) {
    		int g1 = mrgsheet[8][i].indexOf("G");
    		String item_mrg = mrgsheet[2][i] + mrgsheet[3][i] + mrgsheet[5][i] + mrgsheet[8][i].substring(0, g1 + 2);
    		for (int j = 0; j < orgsheet[0].length; j++) {
    			int g2 = orgsheet[8][j].indexOf("G");
    			String item_org = orgsheet[2][j] + orgsheet[3][j] + orgsheet[5][j] + orgsheet[8][j].substring(0, g2 + 2);
    			if (item_org.equals(item_mrg)) {
    				ver[i] = orgver;
    				break;
    			}
    		}
    	}
    	
    	return ver;
	}
    
    // merge 2 version-sheet into one
    private static String[] mergeVer (String[] ver1, String[] ver2) {
    	String[] mrgver = new String [ver1.length];
    	for (int i = 0; i < ver1.length; i++) {
    		mrgver[i] = ver1[i] + "|" + ver2[i];
    		int n = mrgver[i].indexOf("null");
    		if (n != -1) {
    			mrgver[i] = mrgver[i].substring(0, n) + mrgver[i].substring(n + 4);
    			int bar = mrgver[i].lastIndexOf("|");
    			mrgver[i] = mrgver[i].substring(0, bar) + mrgver[i].substring(bar + 1);
    		}
    	}
    	
    	return mrgver;
    }
	
    // get each row's layer number
    private static String[] getLayerNum (String[][] all) {
		String[] lr_num = new String[all[0].length];
		for (int i = 0; i < all[0].length; i++) {
			if (all[8][i].substring(0, 3).equals("DCI")) {
				lr_num[i] = "1";
			}
			else if (all[8][i].indexOf("G2") != -1) {
				lr_num[i] = "2";
			}
			else if (all[8][i].indexOf("G4") != -1) {
				lr_num[i] = "3";
			}
		}
		
		return lr_num;
	}
    
    // trans merged sheet to ipd format
	private static String[][] getIPD (String[][] all, String[] ver) throws Exception {
		int l1_length = HowManyL1 (all);
		String[] lr_num = getLayerNum (all);
		String[][] ipd = new String[16][all[0].length + l1_length];
		for (int i = 0; i < l1_length; i++) {
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
		for (int i = 0; i < (all[0].length - l1_length); i++) {
			ipd[5][2*l1_length + i] = lr_num[l1_length + i];
			ipd[6][2*l1_length + i] = "blank";
			ipd[7][2*l1_length + i] = all[2][l1_length + i];
			ipd[8][2*l1_length + i] = all[3][l1_length + i];
			ipd[9][2*l1_length + i] = "零件名称";
			ipd[10][2*l1_length + i] = all[4][l1_length + i];
			ipd[14][2*l1_length + i] = ver[l1_length + i];
			ipd[15][2*l1_length + i] = all[5][l1_length + i];
		}

		for (int i = 2*l1_length; i < ipd[0].length; i++) {
			if (ipd[7][i].indexOf("G99") != -1) {
				ipd[6][i] = "highlight";
			}
		}
		
		for (int i = 2*l1_length; i < ipd[0].length; i++) {
			int undrln = ipd[7][i].indexOf("_");
			if (undrln != -1) {
				ipd[7][i] = ipd[7][i].substring(0, undrln) + "/" + ipd[7][i].substring(undrln + 1);
			}
		}
		
		return ipd;
	}
	
	// write the ipd sheet into template
	private static void writeIntoTemplate (String[][] newsheet, String savepath) throws Exception {
		Workbook wb = Workbook.getWorkbook 
	    (new File("src" + System.getProperty("file.separator") + 
	    "resources" + System.getProperty("file.separator") + "template.xls")); 
		WritableWorkbook wwb = Workbook.createWorkbook(new File(savepath), wb);
		WritableSheet ws = wwb.getSheet(0);
		WritableFont font1 = new WritableFont(WritableFont.createFont("Arial"),10, WritableFont.NO_BOLD); // 字体样式
        WritableCellFormat wcf1 = new WritableCellFormat(font1);
        wcf1.setBackground(Colour.YELLOW);
        WritableFont font2 = new WritableFont(WritableFont.createFont("宋体"),10, WritableFont.NO_BOLD); // 字体样式
        WritableCellFormat wcf2 = new WritableCellFormat(font2);
		for (int i = 0; i < newsheet.length; i++) {
			for (int j = 0; j < newsheet[0].length; j++){
				if (newsheet[i][j] != "null") {
					ws.addCell(new Label(i, j + 2, newsheet[i][j], wcf2));					
				}	
			}
		}
		for (int i = 0; i < newsheet[0].length; i++) {
			if (newsheet[7][i].equals("zero") ) {
				ws.addCell(new Label(7, i + 2, ""));					
			}
			else if (newsheet[6][i].equals("highlight")) {
				ws.addCell(new Label(6, i + 2, "", wcf1));
			}
		}
		wwb.write();
		wwb.close();
		wb.close();		
	}
	
	// make output ipd by input multiple sheet
	protected static void makeOutput (String[] path, String savepath) throws Exception {
		int n = path.length;
		List<String[][]> lst = new ArrayList<String[][]> ();
		for (int i = 0; i < n; i++) {
			lst.add(getSheet (path[i]));
		}
		for (int i = 0; i < n; i++) {
			lst.add(getLayer (lst.get(i), 1));
		}
		for (int i = 0; i < n; i++) {
			lst.add(getLayer (lst.get(i), 2));
		}
		for (int i = 0; i < n; i++) {
			lst.add(getLayer (lst.get(i), 3));
		}
		for (int i = 2*n; i < 4*n; i++) {
			lst.set(i, remove_R (lst.get(i)));
		}
		String[][] l1 = mergeL1 (lst.get(n), lst.get(n + 1));
		for (int i = n + 2; i < 2*n; i++) {
			l1 = mergeL1 (l1, lst.get(i));
		}
		String [][] l2 = mergeL2(lst.get(2*n), lst.get(2*n + 1));
		for (int i = 2*n + 2; i < 3*n; i++) {
			l2 = mergeL2 (l2, lst.get(i));
		}
		
		String[][] l3 = mergeL3
		(lst.get(3*n), lst.get(3*n + 1), getL2Num(lst.get(0), lst.get(3*n)), getL2Num(lst.get(1), lst.get(3*n + 1)));
		for (int i = 3*n + 2;i < 4*n; i++) {
			l3 = mergeL3 (l3, lst.get(i), getL2Num2(l3), getL2Num(lst.get(i - 3*n), lst.get(i)));
		}
		
		String[][] l2l3 = insertL2L3(l2, l3);
		
		String[][] all = insertL1(l1, l2l3);
		
		List<String[]> lstver = new ArrayList<String[]> ();
		for (int i = 0; i < n; i++) {
			lstver.add(getVer (all, lst.get(i)));
		}
		String[] ver = mergeVer (lstver.get(0), lstver.get(1));
		for (int i = 2; i < n; i++) {
			ver = mergeVer (ver, lstver.get(i));
		}
		
		String[][] ipd = getIPD(all, ver);
		
		writeIntoTemplate(ipd, savepath); 
		
	}

	


}

