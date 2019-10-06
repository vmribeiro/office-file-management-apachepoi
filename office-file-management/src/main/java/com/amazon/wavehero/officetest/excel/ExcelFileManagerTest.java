package com.amazon.wavehero.officetest.excel;

import java.io.IOException;
import java.util.ArrayList;

/**
 * 
 * Test class
 *
 * @author Victor Ribeiro
 *
 */
public class ExcelFileManagerTest {
	public static void main(String[] args) throws IOException {

		// Instantiate the class
		ExcelFileManager fileManager = new ExcelFileManager();
		
		// Get the sheets that will be read
		ArrayList<Integer> sheetNumberArray = new ArrayList<Integer>();
		sheetNumberArray.add(0);

		// Read the file
		ArrayList<ArrayList<ArrayList<String>>> fileData = fileManager.readXLSXFile("input_files/input_excel_test.xlsx",
				sheetNumberArray);

		// Show the data
		fileManager.printFileData(fileData);
		
		
		// Treat null columns
		int sheetNum = 0, rowNum = 0, colNum = 4;
		
		try{ 
			System.out.println(fileData.get(sheetNum).get(rowNum).get(colNum));
		} catch (IndexOutOfBoundsException e) {
			System.out.println("\nWarning: Sheet " + sheetNum + " row " + rowNum + " column " + colNum + " does not have any values");
		}

	}
}
