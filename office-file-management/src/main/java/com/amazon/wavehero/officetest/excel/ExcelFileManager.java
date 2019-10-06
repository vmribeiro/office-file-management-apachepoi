package com.amazon.wavehero.officetest.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Manage excel files (read and write)
 *
 * @author Victor Ribeiro
 *
 */
public class ExcelFileManager {
	
	private static XSSFWorkbook workbook;
	private static XSSFWorkbook wb;

	/**
	 * Print file data array on the console
	 * @param fileData
	 *
	 * @author Victor Ribeiro
	 *
	 */
	public void printFileData(ArrayList<ArrayList<ArrayList<String>>> fileData) {
		for (ArrayList<ArrayList<String>> sheet : fileData) {
			for (ArrayList<String> row : sheet) {
				for (String cell : row) {
					System.out.print(cell + "\t");
				}
				System.out.println("");
			}
		}
	}

	/**
	 * Reads each one of the indicated sheets (number) in a XLSX file and returns an Array of String ArrayLists
	 * 
	 * @param filePath
	 * @param sheetNumberArray
	 * @return ArrayList<ArrayList<ArrayList<String>>> || The first array (from
	 *         inside to outside) are the rows and the second are the sheets
	 * @throws IOException
	 *
	 * @author Victor Ribeiro
	 *
	 */
	public ArrayList<ArrayList<ArrayList<String>>> readXLSXFile(String filePath,
			ArrayList<Integer> sheetNumberArray) throws IOException {

		// Initialize the array that will be returned
		ArrayList<ArrayList<ArrayList<String>>> sheetsData = new ArrayList<ArrayList<ArrayList<String>>>();

		// Open buffer to read the file
		InputStream ExcelFileToRead = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(ExcelFileToRead);

		// Read each one of the sheets
		for (Integer sheetNumber : sheetNumberArray) {

			// Open the sheet
			XSSFSheet sheet = workbook.getSheetAt(sheetNumber);

			// Initialize the Array that will store the data regarding each row
			ArrayList<ArrayList<String>> sheetData = new ArrayList<ArrayList<String>>();

			// Initialize variables for the next loop
			XSSFRow row;
			XSSFCell cell;

			// Read each row until there is none left
			Iterator rows = sheet.rowIterator();
			while (rows.hasNext()) {

				// Initialize the array that will store the data for the row
				ArrayList<String> rowData = new ArrayList<String>();

				// Increment the row iterator
				row = (XSSFRow) rows.next();

				// Read each cell until there is none left
				Iterator cells = row.cellIterator();
				while (cells.hasNext()) {

					// Increment the cell iterator
					cell = (XSSFCell) cells.next();

					// Action block
					if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {

						// Adds the cell value to the row array
						rowData.add(cell.getStringCellValue().toString());

					} else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {

						// Adds the cell value to the row array
						rowData.add(((Double) cell.getNumericCellValue()).toString());

					} else {
						// Handle other cases
					}

				}

				// Add the row to the sheet array
				sheetData.add(rowData);

			}

			// Add the sheet to the Sheets array
			sheetsData.add(sheetData);

		}

		return sheetsData;

	}

	/**
	 * Writes a XLSX file according to a String matrix parameter.
	 * @param filePath
	 * @param sheetName
	 * @param sheetValuesMatrix
	 * @throws IOException
	 *
	 * @author Victor Ribeiro
	 *
	 */
	public void writeXLSXFile(String filePath, String sheetName, String[][] sheetValuesMatrix)
			throws IOException {

		// Create a new workbook
		wb = new XSSFWorkbook();

		// Create a new sheet
		XSSFSheet sheet = wb.createSheet(sheetName);

		// Write the file

		// # Read the matrix and write it to the file
		// # r -> row || c -> column
		for (int r = 0; r < sheetValuesMatrix.length; r++) {

			// Generate a new row
			XSSFRow row = sheet.createRow(r);

			for (int c = 0; c < sheetValuesMatrix[0].length; c++) {

				// Create a new cell
				XSSFCell cell = row.createCell(c);

				// Fill the cell
				cell.setCellValue(sheetValuesMatrix[r][c]);

			}
		}

		// Resize all columns of the file
		for (int c = 0; c < sheetValuesMatrix[0].length; c++) {
			sheet.autoSizeColumn(c);
		}

		// Open the buffer to write the file in the specified path
		FileOutputStream fileOut = new FileOutputStream(filePath);

		// Write this workbook to an OutputStream
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();

		/**
		 * ******************************* ********FORMATTING TIPS********
		 * *******************************
		 * 
		 * DateFormat format = workbook.createDataFormat(); CellStyle dateStyle
		 * workbook.createCellStyle("yyyy-mm-dd");
		 * dateStyle.setDataFormat(format.getFormat("yyyy-mm-dd"));
		 * cell.setCellStyle(dateStyle); cell.setCellValue(new Date());
		 * 
		 * 
		 * 
		 */
	}

}
