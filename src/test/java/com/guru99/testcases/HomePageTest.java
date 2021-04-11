package com.guru99.testcases;

import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import com.guru99.base.TestBase;
import com.guru99.pages.HomePage;


public class HomePageTest extends TestBase {
	HomePage homePage;

	public HomePageTest() {
		super();
	}

	@BeforeMethod
	public void setUp() {
		initialization();
		homePage = new HomePage();
	}
	
	@Test(priority=1)
	public void printNumberOfTopicsUnderGivenHeadingTest(){
		Assert.assertTrue(homePage.assignmentOne_printNumberOfTopicsUnderGivenHeading());
	}
	
	@Test(priority=2)
	public void searchGivenCourseAndFindRecordCountTest(){
		Assert.assertTrue(homePage.assignmentTwo_searchGivenCourseAndFindRecordCount());
	}
	
	@AfterMethod
	public void tearDown(){
		driver.quit();
	}

}
