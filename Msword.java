package com.msword;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Msword {

	public static void main(String[] args) throws Exception {
		
			
		String temp="";
		
		String fileName="D:\\Users\\lkommine\\Desktop\\Work\\table.docx";
		XWPFDocument doc = new XWPFDocument(new FileInputStream(fileName));
		int tableIdx = 1;
		int rowIdx = 1;
		int colIdx = 1;
		List<XWPFTable> table = doc.getTables();

		System.out.println(
				"==========No Of Tables in Document=============================================" + table.size());

		for (XWPFTable xwpfTable : table) {

			System.out.println("================table -" + tableIdx + "===Data==");

			rowIdx = 1;
			List<XWPFTableRow> row = xwpfTable.getRows();

			System.out.println("total rows------------ " + row.size());

			for (XWPFTableRow xwpfTableRow : row) {

				System.out.println("Row -" + rowIdx);

				colIdx = 1;
				List<XWPFTableCell> cell = xwpfTableRow.getTableCells();
				for (XWPFTableCell xwpfTableCell : cell) {
					
					if (xwpfTableCell != null) {
						
						if(xwpfTableCell.getText().length()>0) {
							
							System.out.print("\t" + colIdx + "- column value: " + xwpfTableCell.getText());
							temp=xwpfTableCell.getText();
						}
						
					}
					
					colIdx++;
					
					if(rowIdx==1 && colIdx==3) {
						if(temp.equals("Name")) {
							xwpfTableCell.setText("Loki");
						}
					}
					
					if(rowIdx==2 && colIdx==3) {
						if(temp.equals("Designation")) {
							xwpfTableCell.setText("Consultant");
						}
					}
					
				}
				System.out.println("");
				rowIdx++;
			}
			tableIdx++;
			System.out.println("");
		}
		OutputStream out = new FileOutputStream(fileName);
		doc.write(out);
		out.close();
			
		
	}

}
