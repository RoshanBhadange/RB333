import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.configuration.RunConfiguration
//import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.WebDriver as WebDriver
import org.openqa.selenium.chrome.ChromeDriver as ChromeDriver
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory
import org.openqa.selenium.By as By
import org.openqa.selenium.By.ByCssSelector as ByCssSelector
/*Change*/ import java.util.List as List
import org.openqa.selenium.WebElement as WebElement
import java.util.ArrayList as ArrayList
import java.lang.String as String
import java.sql.ResultSet as ResultSet
import org.openqa.selenium.interactions.Action as Action
import org.openqa.selenium.interactions.Actions as Actions
import com.kms.katalon.core.logging.KeywordLogger as KeywordLogger
import com.kms.katalon.core.annotation.Keyword as Keyword
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import org.apache.poi.ss.usermodel.Sheet
WebUI.callTestCase(findTestCase('test_cases_repo/MSTR_Login'), [:], FailureHandling.STOP_ON_FAILURE)

WebUI.maximizeWindow()

WebDriver driver = DriverFactory.getWebDriver()

/**   
 * Reading MSTR table values
 * Reading row data
 */
String rowpath= ""
String filePath= "D://webData.xlsx"
String WebTableSheetName= "RegionalMSTRData"
String DBTableSheetName= "RegionalSalesDBData"
//CustomKeywords.'com.infocepts.writeDataToExcel.OutPutData.writeWebTableData'(rowpath,filePath,WebTableSheetName)


for (int i =0;i<=6;i++)
{
	rowpath = "//*[@id='mstr105']/table/tbody/tr" 
   
	CustomKeywords.'com.infocepts.writeDataToExcel.OutPutData.writeWebTableData'(rowpath,filePath,WebTableSheetName)
	
}
/**
 * Reading DB table values
 */


String Query = null

for (def rowNum = 1; rowNum <= findTestData('database_data/READQUERY').getRowNumbers(); rowNum++) {
	Query = findTestData('database_data/READQUERY').getValue(3, rowNum)

	println(Query)
}

String dbType = "thin"
CustomKeywords.'com.infocepts.database.dbconnect.connectDB'(dbType,GlobalVariable.dbUrl, GlobalVariable.dbName, GlobalVariable.dbPort, GlobalVariable.dbUser,'root')
ResultSet resultSet = CustomKeywords.'com.infocepts.database.dbconnect.executeQuery'(Query)
/**
 * Writing DB data in excel
 */
CustomKeywords.'com.infocepts.writeDataToExcel.OutPutData.writeDBData'(resultSet, filePath, DBTableSheetName)
CustomKeywords.'com.infocepts.database.dbconnect.closeDatabaseConnection'()

/**
 *Validating sheets 
 */
Sheet webTableSheet = CustomKeywords.'com.infocepts.writeDataToExcel.OutPutData.getExcelSheetByName'(filePath,WebTableSheetName)
Sheet DbTableSheet = CustomKeywords.'com.infocepts.writeDataToExcel.OutPutData.getExcelSheetByName'(filePath,DBTableSheetName)
CustomKeywords.'com.infocepts.writeDataToExcel.OutPutData.compareTwoSheets'(webTableSheet, DbTableSheet,filePath,"Result")

WebUI.closeBrowser() 


