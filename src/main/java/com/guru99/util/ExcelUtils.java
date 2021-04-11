package com.guru99.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	private XSSFSheet xlWSheet;
	private XSSFWorkbook xlWBook;
	private XSSFRow row;
	private XSSFCell cell;
	private DataFormatter formatter = new DataFormatter();
	public int intHeader = 0;
	private static String testDataFile = "";
	
	public ExcelUtils(String path) throws Exception {
		try {
//			FileInputStream excelFile = new FileInputStream(path);
//			OPCPackage excelFile = OPCPackage.open(path);
//			xlWBook = new XSSFWorkbook(excelFile);
			
			setExcelFile(path);
		} catch (Exception e) {
			throw new Exception(e);
		}
	}
	
	
	private void setExcelFile(String Path) throws Exception {
		try {
			FileInputStream excelFile = new FileInputStream(Path);
			xlWBook = new XSSFWorkbook(excelFile);
			testDataFile = Path;
		} catch (Exception e) {
			throw new Exception(e);
		}
	}
	

	public XSSFWorkbook getWorkBook() throws Exception{
		try {
			return xlWBook;
			
		}catch(Exception e) {
			throw new Exception(e);
		}
	}

	public void setWorkBook(XSSFWorkbook wkBook) throws Exception{
		try {
			xlWBook = wkBook;
		}catch(Exception e) {
			throw new Exception(e);
		}
	}

	public XSSFSheet getWorkSheet(String sheetName) throws Exception{
		try {
			xlWSheet = xlWBook.getSheet(sheetName);
			return xlWSheet;
		}catch(Exception e) {
			throw new Exception(e);
		}
	}

	public void setWorkSheet(XSSFSheet wkSheet) throws Exception{
		try {
			xlWSheet = wkSheet;
		}catch(Exception e) {
			throw new Exception(e);
		}
		
	}

	
	public String getMappedCellData(int RowNum, String ColNum, String SheetName, int rowTDS, String sheetTestData) throws Exception {
		try {
			xlWSheet = xlWBook.getSheet(SheetName);
			cell = xlWSheet.getRow(RowNum).getCell(Integer.parseInt(ColNum));
			String CellData = cell.getStringCellValue();
			if(CellData.contains("Key:")) {
				String key = CellData.split(":")[1]; 
				int columnIndex = getColumnIndex(key, sheetTestData);
				CellData = getCellData(rowTDS, columnIndex, sheetTestData);
			}
			return CellData;
		} catch (Exception e) {
			throw new Exception(e);
		}
	}

	
	public String getCellData(int RowNum, String ColName, XSSFSheet xlWSheet) throws Exception {
		try {
			cell = xlWSheet.getRow(RowNum).getCell(getColumnIndex(ColName, xlWSheet));
			String CellData = cell.getStringCellValue();
			return CellData;
		} catch (Exception e) {
			throw new Exception(e);
		}
	}

	
	public int getColumnIndex(String ColName, XSSFSheet xlWSheet) throws Exception {
		int colNum = 0;
		try {
			row = xlWSheet.getRow(0);
		    for(;colNum < row.getLastCellNum(); colNum++) {
		    	if(getCellData(0, colNum, xlWSheet).equalsIgnoreCase(ColName)) {
		    		break;
		    	}
		    }
			return colNum;
		} catch (Exception e) {
			throw new Exception(e);
		}
	}

	
	public String getCellData(int RowNum, int ColNum, XSSFSheet xlWSheet) throws Exception {
		try {
			cell = xlWSheet.getRow(RowNum).getCell(ColNum);
			String CellData = cell.getStringCellValue();
			return CellData;
		} catch (Exception e) {
			throw new Exception(e);
		}
	}

	
	public int getRowCount(XSSFSheet xlWSheet) throws Exception {
		int iNumber = 0;
		try {
			iNumber = xlWSheet.getLastRowNum() + 1;
			return iNumber;
		} catch (Exception e) {
			throw new Exception(e);
		}
		
	}

	
	public String getCellData(int RowNum, String ColName, String SheetName) throws Exception {
		try {
			xlWSheet = xlWBook.getSheet(SheetName);
			cell = xlWSheet.getRow(RowNum).getCell(getColumnIndex(ColName, SheetName));
			String CellData = formatter.formatCellValue(cell);
			return CellData;
		} catch (Exception e) {
			throw new Exception(e);
		}
	}

	
	public String getCellData(int RowNum, int ColNum, String SheetName) throws Exception {
		try {
			xlWSheet = xlWBook.getSheet(SheetName);
			cell = xlWSheet.getRow(RowNum).getCell(ColNum);
			String CellData = formatter.formatCellValue(cell);
			return CellData;
		} catch (Exception e) {
			throw new Exception(e);
		}
	}


	public int getColumnIndex(String ColName, String SheetName) throws Exception {
		int colNum = 0;
		int rowNum = intHeader;
		try {
			xlWSheet = xlWBook.getSheet(SheetName);
			row = xlWSheet.getRow(rowNum);
		    for(;colNum < row.getLastCellNum(); colNum++) {
		    	if(getCellData(rowNum, colNum, SheetName).equalsIgnoreCase(ColName)) {
		    		break;
		    	}
		    }
			return colNum;
		} catch (Exception e) {
			throw new Exception(e);
		}
	}
	


	public int getRowCount(String SheetName) throws Exception {
		int iNumber = 0;
		try {
			xlWSheet = xlWBook.getSheet(SheetName);
			iNumber = xlWSheet.getLastRowNum() + 1;
			return iNumber;
		}catch(NullPointerException npe) {
			npe.printStackTrace();
			throw new Exception(npe);
		} catch (Exception e) {
			throw new Exception(e);
		}
		
	}

	public int getRowContains(String keyToLookFor, String colNum, XSSFSheet xlWSheet) throws Exception {
		int iRowNum = 0;
		try {
			int rowCount = getRowCount(xlWSheet);
			for (; iRowNum < rowCount; iRowNum++) {
				if (getCellData(iRowNum, colNum, xlWSheet).equalsIgnoreCase(keyToLookFor)) {
					break;
				}
			}
		} catch (Exception e) {
			throw new Exception(e);
		}
		return iRowNum;
	}

	public void setCellData(String Result, int RowNum, String ColNum, XSSFSheet xlWSheet) throws Exception {
		try {
			row = xlWSheet.getRow(RowNum);
			cell = row.getCell(Integer.parseInt(ColNum), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
			if (cell == null) {
				cell = row.createCell(Integer.parseInt(ColNum));
				cell.setCellValue(Result);
			} else {
				cell.setCellValue(Result);
			}
			FileOutputStream fileOut = new FileOutputStream(testDataFile);
			xlWBook.write(fileOut);
			fileOut.close();
//			xlWBook = new XSSFWorkbook(new FileInputStream(testDataFile));
		} catch (Exception e) {
			throw new Exception(e);

		}
	}

	
	public int getRowContains(String keyToLookFor, String colNum, String SheetName) throws Exception {
		int iRowNum = 0;
		try {
			int rowCount = getRowCount(SheetName);
			for (iRowNum = intHeader; iRowNum < rowCount; iRowNum++) {
				if (getCellData(iRowNum, colNum, SheetName).equalsIgnoreCase(keyToLookFor)) {
					break;
				}
			}
		} catch (Exception e) {
			throw new Exception(e);
		}
		return iRowNum;
	}

	
	public void setCellData(String Result, int RowNum, String ColNum, String SheetName) throws Exception {
		try {

			xlWSheet = xlWBook.getSheet(SheetName);
			row = xlWSheet.getRow(RowNum);
			cell = row.getCell(Integer.parseInt(ColNum), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);			
			if (cell == null) {
				cell = row.createCell(Integer.parseInt(ColNum));
				cell.setCellValue(Result);
			} else {
				cell.setCellValue(Result);
			}
			FileOutputStream fileOut = new FileOutputStream(testDataFile);
			xlWBook.write(fileOut);
			fileOut.close();
			xlWBook = new XSSFWorkbook(new FileInputStream(testDataFile));
		} catch (Exception e) {
			throw new Exception(e);
		}
	}
}
