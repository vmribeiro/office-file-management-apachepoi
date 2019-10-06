package com.officemanagement.excel;

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

		// ********************
		// Read
		// ********************

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

		try {
			System.out.println(fileData.get(sheetNum).get(rowNum).get(colNum));
		} catch (IndexOutOfBoundsException e) {
			System.out.println("\nWarning: Sheet " + sheetNum + " row " + rowNum + " column " + colNum
					+ " does not have any values");
		}

		// ********************
		// Write
		// ********************

		// Get the highest column number
		int maxColSize = 0;
		for (ArrayList<String> row : fileData.get(sheetNum)) { // fileData.get(sheetNum) -> for the sheet 0 it will search in each row (ArrayList<String>) the highest size of the column
			if (row.size() > maxColSize) {
				maxColSize = row.size();
			}
		}

		// Create the matrix
		String[][] sheetValuesMatrix = new String[fileData.get(sheetNum).size()][maxColSize];

		// Fill the matrix for the Sheet 0
		int r = 0, c = 0; // r -> row || c -> column/cell
		for (ArrayList<String> row : fileData.get(sheetNum)) {

			for (String cell : row) {

				sheetValuesMatrix[r][c] = cell;

				c++;
			}

			c = 0;
			r++;
		}

		fileManager.writeXLSXFile("output_files/excel-test-array.xlsx", "Sheet1", sheetValuesMatrix);

	}
}
