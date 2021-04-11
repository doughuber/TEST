package com.guru99.pages;

import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.guru99.util.ExcelUtils;
import com.guru99.base.TestBase;


public class HomePage extends TestBase
{
	By lnkHomePage = By.xpath("//a[@title='Guru99']");
	By lblTesting = By.xpath("(//h4[*/text()='Testing']/following-sibling::ul)[1]//a");
	By lblMustLearn = By.xpath("(//h4[*/text()='Must Learn!']/following-sibling::ul)[1]//a");
	By txtSearchCourse = By.xpath("(//input[@name='search'])[2]");
	By btnSearchCourse = By.xpath("(//button[*//text()='search'])[2]");
	By lblSearchResultsInfo = By.xpath("(//div[contains(@class,'result-info')])[2]");
	By btnCloseResultsPopUp = By.xpath("(//div[contains(@class,'results-close-btn')])[2]");
	
	
    public boolean assignmentOne_printNumberOfTopicsUnderGivenHeading()
    {
    	List<WebElement> topics = null;
        try {
        	WebDriverWait wait = new WebDriverWait(driver, TIME_OUT_IN_SECONDS);
        	wait.until(ExpectedConditions.visibilityOfElementLocated(lnkHomePage));
        	
        	//	Find all the sub topics under [Testing]
			topics = driver.findElements(lblTesting);
			System.out.println("Number of topics under ['Testing'] are : " + topics.size());
	         
        	//	Find all the sub topics under [Must Learn!]
			topics = driver.findElements(lblMustLearn);
			System.out.println("Number of topics under ['Must Learn!'] are : " + topics.size());
			
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
    }

    
    public boolean assignmentTwo_searchGivenCourseAndFindRecordCount()
    {
        try {
	        ExcelUtils xlWBook = new ExcelUtils(System.getProperty("user.dir") + 
	        		"/src/main/java/com/guru99/testdata/" + "TestData.xlsx");

	        //	Create worksheet reference for sheet - [Course Details]
	        XSSFSheet xlWSheet = xlWBook.getWorkSheet("Course Details");
	        
	        //	Get total number of rows in sheet - [Course Details]
	        int rowCount = xlWBook.getRowCount(xlWSheet);
	        
	        //	Get Column Index for column - [Results Count] for sheet - [Course Details]
	        int resultColumnIndex = xlWBook.getColumnIndex("Results Count", xlWSheet);
	        
	        //	Iterate through sheet - [Course Details] for all rows
	        for(int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
	        	//	Read course topic from test data sheet
	        	String strCourse = xlWBook.getCellData(rowIndex, "Course Topics", xlWSheet);
	        	System.out.println("Searching Course : " + strCourse);
	        	
	        	//	Enter course topic into Search text field
	        	driver.findElement(txtSearchCourse).sendKeys(strCourse);
	        	
	        	//	click on Search button
	        	driver.findElement(btnSearchCourse).click();

	        	//	Create reference for WebDriverWait class
	        	WebDriverWait wait = new WebDriverWait(driver, TIME_OUT_IN_SECONDS);
	        	
	        	//	Wait until Results are loaded
	        	wait.until(ExpectedConditions.visibilityOfElementLocated(lblSearchResultsInfo));

	        	//	Extract search result information
	        	String searchText = driver.findElement(lblSearchResultsInfo).getText();
	        	
	        	//	Split result info string to get the numeric part
	        	String resultCount = searchText.split("results")[0].replace("About ", "");
	        	System.out.println("Results found for course [" + strCourse + "] is : " + resultCount);
	        	
	        	Thread.sleep(1000);
	        	
	        	//	Update the test data sheet with the record count extracted
	        	xlWBook.setCellData(resultCount, rowIndex, String.valueOf(resultColumnIndex), xlWSheet);
	        	
	        	Thread.sleep(1000);
	        	
	        	//	Close the result window
	        	driver.findElement(btnCloseResultsPopUp).click();
	        }
        	return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
    }
}
