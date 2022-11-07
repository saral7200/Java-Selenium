package com.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	public static WebDriver driver;
//1
	public void getChrome() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
	}
	public void getEdge() {
		WebDriverManager.edgedriver().setup();
		driver = new EdgeDriver();
	}
//2
	public void enterApplnUrl(String Url) {
		driver.get(Url);
	}
//3	
	public void maximizeWindow () {
		driver.manage().window().maximize();
	}
//4	
	public void elementSendKeys(WebElement element, String data) {
		element.sendKeys(data);
	}
//5
	public void elementSendKeysJs(WebElement element, String data) {
		JavascriptExecutor jse = (JavascriptExecutor)driver;
		jse.executeScript("arguments[0].setAttribute('value','"+data+"')", element);
	}
//6
	public void elementClick(WebElement element) {
		element.click();
	}
//7
	public WebElement findElementById(String attributeValue) {
		WebElement element = driver.findElement(By.id(attributeValue));
		return element;
	}
//8	
	public WebElement findElementByName(String attributeValue) {
		WebElement element = driver.findElement(By.name(attributeValue));
		return element;
	}
//9
	public WebElement findElementByClassName(String attributeValue) {
		WebElement element = driver.findElement(By.className(attributeValue));
		return element;
	}
//10	
	public WebElement findElementByXpath(String xpath) {
		WebElement element = driver.findElement(By.xpath(xpath));
		return element;
	}
//11
	public void closeWindow() {
		driver.close();
	}
//12	
	public void quitWindow() {
		driver.quit();
	}
//13
	public String getApplnTitle() {
		String title = driver.getTitle();
		return title;
	}
//14
	public String getCurrentUrl() {
		String text = driver.getCurrentUrl();
		return text;
	}
//15
	public String elementGetText(WebElement element) {
		String text = element.getText();
		return text;
	}
//16
	public String elementGetAttributeValue(WebElement element) {
		String attribute = element.getAttribute("value");
		return attribute;
	}
//17
	public String elementGetAttributeValue(WebElement element, String attributename) {
		String attribute = element.getAttribute(attributename);
		return attribute;
	}
//18
	public void selectOptionsByText(WebElement element, String text) {
		Select select = new Select(element);
		select.selectByVisibleText(text);
	}
//19
	public void selectOptionsByAttribute(WebElement element, String value) {
		Select select = new Select(element);
		select.selectByValue(value);
	}
//20
	public void selectOptionsByIndex(WebElement element, int index) {
		Select select = new Select(element);
		select.selectByIndex(index);
	}
//1	
	public void elementClickJs(WebElement element) {
		JavascriptExecutor jse = (JavascriptExecutor)driver;
		jse.executeScript("arguments[0].click()", element);
	}
//2
	public void srollDown(WebElement element) {
		JavascriptExecutor jse = (JavascriptExecutor)driver;
		jse.executeScript("arguments[0].scrollIntoView(true)", element);
	}
//3
	public void srollUp(WebElement element) {
		JavascriptExecutor jse = (JavascriptExecutor)driver;
		jse.executeScript("arguments[0].scrollIntoView(false)", element);
	}
//4
	public boolean dropdownIsMultiple(WebElement element) {
		Select select = new Select(element);
		boolean multiple = select.isMultiple();
		return multiple;
	}
//5
	public WebElement getFirstSelectedOption(WebElement element) {
		Select select = new Select(element);
		WebElement firstSelectedOption = select.getFirstSelectedOption();
		return firstSelectedOption;
	}
//6
	public List<WebElement> getAllSelectedOptions(WebElement element) {
		Select select = new Select(element);
		List<WebElement> allSelectedOptions = select.getAllSelectedOptions();
		return allSelectedOptions;
	}
//7
	public void deSelectByIndex(WebElement element, int index) {
		Select select = new Select(element);
		select.deselectByIndex(index);
	}
//8
	public void deSelectAttribute(WebElement element, String value) {
		Select select = new Select(element);
		select.deselectByValue(value);
	}
//9
	public void deSelectText(WebElement element, String text) {
		Select select = new Select(element);
		select.deselectByVisibleText(text);
	}
//10
	public void deSelectAll(WebElement element) {
		Select select = new Select(element);
		select.deselectAll();
	}
//11
	public void toMoveCursor(WebElement element) {
		Actions ac = new Actions(driver);
		ac.moveToElement(element).perform();
	}
//12
	public void doubleClick(WebElement element) {
		Actions ac = new Actions(driver);
		ac.doubleClick(element).perform();
	}
//13
	public void rightClick(WebElement element) {
		Actions ac = new Actions(driver);
		ac.contextClick(element).perform();
	}
//13
	public void navigateToUrl(String Url) {
		driver.navigate().to(Url);
	}
//14
	public void refreshPage() {
		driver.navigate().refresh();
	}
//15
	public void previousPage() {
		driver.navigate().back();
	}
//16
	public void nextPage() {
		driver.navigate().forward();
	}
//16
	public void clearTextBox(WebElement element) {
		element.clear();
	}
//17
	public boolean isDisplayed(WebElement element) {
		boolean displayed = element.isDisplayed();
		return displayed;
	}
//18
	public boolean isEnabled(WebElement element) {
		boolean enabled = element.isEnabled();
		return enabled;
	}
//19
	public boolean isSelected(WebElement element) {
		boolean selected = element.isSelected();
		return selected;
	}
//20
	public List<WebElement> getAllOptionsFromDropDown(WebElement element) {
		Select select = new Select(element);
		List<WebElement> options = select.getOptions();
		return options;
	}
//21
	public File takeScreenShot() {
		TakesScreenshot ts = (TakesScreenshot)driver;
		File screenshotAs = ts.getScreenshotAs(OutputType.FILE);
		return screenshotAs;
	}
//22
	public void okAlert() {
		Alert at = driver.switchTo().alert();
		at.accept();
	}
//23
	public void cancelAlert() {
		Alert at = driver.switchTo().alert();
		at.dismiss();
	}
//24
	public void sendKeyAlert(String value) {
		Alert at = driver.switchTo().alert();
		at.sendKeys(value);
	}
//25
	public String getTextAlert() {
		Alert at = driver.switchTo().alert();
		String text = at.getText();
		return text;
	}
//26
	public void switchToParentFrame(String id) {
		driver.switchTo().parentFrame();
	}
//27
	public void switchToFrameById(String id) {
		driver.switchTo().frame(id);
	}
//28
	public void switchToFrameByWebElementRef(WebElement element) {
		driver.switchTo().frame(element);
	}
//29
	public void switchToFrameByIndex(int index) {
		driver.switchTo().frame(index);
	}
//30
	public void switchToWindowByUrl(String Url) {
		driver.switchTo().window(Url);
	}
//31
	public void switchToWindowById(String id) {
		driver.switchTo().window(id);
	}
//32
	public void switchToWindowByTitle(String value) {
		driver.switchTo().window(value);
	}
//33
	public void switchToParentWindow() {
		driver.switchTo().defaultContent();
	}
//34
	public String getParentWindow() {
		String parentWindow = driver.getWindowHandle();
		return parentWindow;
	}
//35
	public Set<String> getAllWindows() {
		Set<String> allWindows = driver.getWindowHandles();
		return allWindows;
	}
//36	
	public void unconditionalWait(int msecs) throws InterruptedException {
		Thread.sleep(msecs);
	}
//37
	public String getDataFromCell(String sheetName, int rownum, int cellnum) throws IOException {
		String res = "";
		File file = new File("C:\\Users\\SARAL VASANTH\\eclipse-workspace\\FrameWork\\Excel\\FrameWork.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(rownum);
		Cell cell = row.getCell(cellnum);
		CellType type = cell.getCellType();
		switch (type) {
		case STRING:
			res = cell.getStringCellValue();
			break;
		case NUMERIC:
			if(DateUtil.isCellDateFormatted(cell)) {
				Date dateCellValue = cell.getDateCellValue();
				SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
				res = dateFormat.format(dateCellValue);
			}
			else {
				double numericCellValue = cell.getNumericCellValue();
				BigDecimal valueOf = BigDecimal.valueOf(numericCellValue);
				res = valueOf.toString();
			}
			break;
		default:
			break;
		}
		return res;
	}
//38
	public void insertCellData(String sheetName, int rownum, int cellnum, String oldData, String newData) throws IOException {
		File file = new File("C:\\Users\\SARAL VASANTH\\eclipse-workspace\\FrameWork\\Excel\\FrameWork.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(rownum);
		Cell cell = row.getCell(cellnum);
		String value = cell.getStringCellValue();
		if(value.equals(oldData)) {
			cell.setCellValue(newData);
		}
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
	}
//39
	public void createCellInExistingRow(String sheetName, int rownum, int cellnum, String newData) throws IOException {
		File file = new File("C:\\Users\\SARAL VASANTH\\eclipse-workspace\\FrameWork\\Excel\\FrameWork.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(rownum);
		Cell cell = row.createCell(cellnum);
		cell.setCellValue(newData);
		FileOutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);
	}
//40
	public void createCell(String sheetName, int rownum, int cellnum, String newData) throws IOException {
		File file = new File("C:\\Users\\SARAL VASANTH\\eclipse-workspace\\FrameWork\\Excel\\FrameWork.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.createRow(rownum);
		Cell cell = row.createCell(cellnum);
		cell.setCellValue(newData);
		FileOutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);
	}
//41
	public void createNewExcelFile(String newPath, String sheetName, int rownum, int cellnum, String data) throws IOException {
		File file = new File(newPath);
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.createRow(rownum);
		Cell cell = row.createCell(cellnum);
		cell.setCellValue(data);
		FileOutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);
	}
//42
	public int rowsCount(String sheetName) throws IOException {
		File file = new File("C:\\Users\\SARAL VASANTH\\eclipse-workspace\\FrameWork\\Excel\\FrameWork.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		int rowCount = sheet.getPhysicalNumberOfRows();
		return rowCount;
	}
//43
	public int cellsCount(String sheetName, int rownum) throws IOException {
		File file = new File("C:\\Users\\SARAL VASANTH\\eclipse-workspace\\FrameWork\\Excel\\FrameWork.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(rownum);
		int cellsCount = row.getPhysicalNumberOfCells();
		return cellsCount;
	}
//44

	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
