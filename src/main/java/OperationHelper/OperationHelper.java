package OperationHelper;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;
import javax.sound.midi.Soundbank;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import DataProvider.DataProvider;
import ReaProperties.ReadProperties;
import cucumber.api.DataTable;
import gherkin.deps.net.iharder.Base64;
import io.github.bonigarcia.wdm.WebDriverManager;
import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;
import ru.yandex.qatools.ashot.shooting.ShootingStrategies;

public class OperationHelper {
	public static WebDriver driver;
	public static DataProvider data = new DataProvider();
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;
	public static XSSFCell cell = null;
	public static FileOutputStream fileOut;
	public static FileInputStream file;

	/**
	 * Return By id,name, tagname, xpath, css, classname, linktext, paralinktext
	 * 
	 * @param objectName
	 * @return
	 * @throws Exception
	 */
	private By getIdentifier(String objectName) throws Exception {
		if (null != ReadProperties.getElementsAndUrls(objectName + "-id")) {
			return By.id(ReadProperties.getElementsAndUrls(objectName + "-id"));
		}
		if (null != ReadProperties.getElementsAndUrls(objectName + "-xp")) {
			return By.xpath(ReadProperties.getElementsAndUrls(objectName + "-xp"));
		}
		if (null != ReadProperties.getElementsAndUrls(objectName + "-css")) {
			return By.cssSelector(ReadProperties.getElementsAndUrls(objectName + "-css"));
		}
		if (null != ReadProperties.getElementsAndUrls(objectName + "-cln")) {
			return By.className(ReadProperties.getElementsAndUrls(objectName + "-cln"));
		}
		if (null != ReadProperties.getElementsAndUrls(objectName + "-link")) {
			return By.linkText(ReadProperties.getElementsAndUrls(objectName + "-link"));
		}
		if (null != ReadProperties.getElementsAndUrls(objectName + "-plink")) {
			return By.partialLinkText(ReadProperties.getElementsAndUrls(objectName + "-plink"));
		}
		if (null != ReadProperties.getElementsAndUrls(objectName + "-tag")) {
			return By.tagName(ReadProperties.getElementsAndUrls(objectName + "-tag"));
		}
		if (null != ReadProperties.getElementsAndUrls(objectName + "-name")) {
			return By.name(ReadProperties.getElementsAndUrls(objectName + "-name"));
		}
		try {
			if (driver.findElement(By.linkText(objectName)).isDisplayed()) {
				return By.linkText(objectName);
			} else {
				return null;
			}
		} catch (Exception e) {
			return null;
		}
	}

	@SuppressWarnings("unused")
	private WebElement getElement(String objectName) throws Exception {
		if (null != ReadProperties.getElementsAndUrls(objectName + "-id")) {
			return driver.findElement(By.id(ReadProperties.getElementsAndUrls(objectName + "-id")));
		} else if (null != ReadProperties.getElementsAndUrls(objectName + "-xp")) {
			return driver.findElement(By.xpath(ReadProperties.getElementsAndUrls(objectName + "-xp")));
		} else if (null != ReadProperties.getElementsAndUrls(objectName + "-css")) {
			return driver.findElement(By.cssSelector(ReadProperties.getElementsAndUrls(objectName + "-css")));
		} else if (null != ReadProperties.getElementsAndUrls(objectName + "-cln")) {
			return driver.findElement(By.className(ReadProperties.getElementsAndUrls(objectName + "-cln")));
		} else if (null != ReadProperties.getElementsAndUrls(objectName + "-link")) {
			return driver.findElement(By.linkText(ReadProperties.getElementsAndUrls(objectName + "-link")));
		} else if (null != ReadProperties.getElementsAndUrls(objectName + "-plink")) {
			return driver.findElement(By.partialLinkText(ReadProperties.getElementsAndUrls(objectName + "-plink")));
		} else if (null != ReadProperties.getElementsAndUrls(objectName + "-tag")) {
			return driver.findElement(By.tagName(ReadProperties.getElementsAndUrls(objectName + "-tag")));
		} else if (null != ReadProperties.getElementsAndUrls(objectName + "-name")) {
			return driver.findElement(By.name(ReadProperties.getElementsAndUrls(objectName + "-name")));
		} else {
			try {
				if (driver.findElement(By.linkText(objectName)).isDisplayed()) {
					return driver.findElement(By.linkText(objectName));
				} else {
					return null;
				}
			} catch (Exception e) {
				return null;
			}
		}
	}

	/* Method wait Element in case lodaing so slowly */
	public static void waitForAjax() throws InterruptedException {
		try {
			boolean statusReturn = false;
			boolean statusReturn1 = false;
			while (true) // Handle timeout somewhere
			{
				JavascriptExecutor executor = (JavascriptExecutor) driver;
				boolean ajaxIsComplete = (Boolean) executor.executeScript("return jQuery.active == 0");
				Long ajaxState = (Long) executor.executeScript("return jQuery.active");
				String readyState = (String) executor.executeScript("return document.readyState");

				if (null == ajaxState) {
					statusReturn = true;
				} else if (null != ajaxState && ajaxIsComplete) {
					statusReturn = true;
				} else {
					statusReturn = false;
				}
				if (null == readyState) {
					statusReturn1 = true;
				} else if (null != readyState && readyState.equals("complete")) {
					statusReturn1 = true;
				} else {
					statusReturn1 = false;
				}

				if (statusReturn && statusReturn1) {
					Thread.sleep(1000);
					break;
				}
				Thread.sleep(100);
			}
		} catch (Exception e) {
			waitForWebState();
		}
	}

	public static void waitForWebState() throws InterruptedException {
		while (true) // Handle timeout somewhere
		{
			JavascriptExecutor js = (JavascriptExecutor) driver;
			String readyState = (String) js.executeScript("return document.readyState");
			if (readyState.equals("complete")) {
				Thread.sleep(1000);
				break;
			}
			Thread.sleep(100);
		}
	}

	public void OpenBrowser(String urls) throws IOException, InterruptedException {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().deleteAllCookies();
		driver.get(urls);
		waitForAjax();
	}

	public static void OpenBrowser() throws IOException, InterruptedException {
		String keyLocalName = ReadProperties.getElementsAndUrls("browser");
		switch (keyLocalName) {
		case "chrome":
			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.manage().deleteAllCookies();
			waitForAjax();
			break;
		default:
			WebDriverManager.iedriver().setup();
			driver = new InternetExplorerDriver();
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.manage().deleteAllCookies();
			waitForAjax();
			break;
		}

	}

	public void getLocalName(String keyLocalName, String urls) {
		switch (keyLocalName) {
		case "chrome":
			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.manage().deleteAllCookies();
			driver.get(urls);
			break;

		default:
			WebDriverManager.chromedriver().setup();
			driver = new InternetExplorerDriver();
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.manage().deleteAllCookies();
			driver.get(urls);
			break;
		}
	}

	public void CheckUsercaseOnExcel(String fileName, String sheetName, String numberRound, DataTable data)
			throws Exception {
		List<List<String>> list = data.raw();
		String rowDo = list.get(0).get(0).trim();
		DataProvider dataPro = new DataProvider();
		// Read file excel
		String strFileName = dataPro.getExcelfile(fileName);
		file = new FileInputStream(strFileName);
		workbook = new XSSFWorkbook(file);
		sheet = workbook.getSheet(sheetName);
		int desInex = 0, expetedInex = 0, resultInex = 0;
		int columnCount = this.getCellCount(strFileName, sheetName);
		int roundResult = Integer.parseInt(numberRound);
		for (int col = 0; col < columnCount; col++) {
			String titleColumnText = sheet.getRow(0).getCell(col).toString().trim();
			if (titleColumnText.equalsIgnoreCase("Description")) {
				desInex = col;
			}
			if (titleColumnText.equalsIgnoreCase("Expected Result")) {
				expetedInex = col;
			}
			if (titleColumnText.equalsIgnoreCase("Result Round " + roundResult)) {
				resultInex = col;
			}
		}
		if (rowDo.equals("All")) {
			int rowCount = sheet.getLastRowNum();
			for (int i = 1; i < rowCount; i++) {
				stepByStep(strFileName, i, desInex, expetedInex, resultInex);
			}
		} else {
			int rowCount = sheet.getLastRowNum();
			if (rowDo.contains(",")) { // Case 2: Run some row on Excel file
				String[] arrRowDo = rowDo.split(",");
				for (int i = 0; i < arrRowDo.length; i++) {
					int row = Integer.parseInt(arrRowDo[i]);
					if (row > rowCount) {
						Assert.assertTrue("Index of row do out of range ", false);
					} else {
						stepByStep(strFileName, row, desInex, expetedInex, resultInex);
					}
				}
			}
		}
	}

	public void CheckResponseCodeUrl(String fileName, String sheetName, String numberRound, DataTable data)
			throws Exception {
		List<List<String>> list = data.raw();
		String rowDo = list.get(0).get(0).trim();
		DataProvider dataPro = new DataProvider();
		// Read file excel
		String strFileName = dataPro.getExcelfile(fileName);
		file = new FileInputStream(strFileName);
		workbook = new XSSFWorkbook(file);
		sheet = workbook.getSheet(sheetName);
		int desInex = 0;
		int columnCount = this.getCellCount(strFileName, sheetName);
		int roundResult = Integer.parseInt(numberRound);
		for (int col = 0; col < columnCount; col++) {
			String titleColumnText = sheet.getRow(0).getCell(col).toString().trim();
			if (titleColumnText.equalsIgnoreCase("Description")) {
				desInex = col;
			}
			if (titleColumnText.equalsIgnoreCase("Result Round " + roundResult)) {
			}
		}
		if (rowDo.equals("All")) {
			int rowCount = sheet.getLastRowNum();
			System.out.println("RowCount" + rowCount);
			for (int i = 1; i < rowCount; i++) {
				String url = sheet.getRow(i).getCell(desInex).getStringCellValue();
				System.out.println(url);
				System.out.println(getResponse(url));
			}
		}
	}

	private int getCellCount(String fileName, String sheetName) throws IOException, Exception {
		FileInputStream fis = new FileInputStream(fileName);
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet(sheetName);
		int number = sheet.getRow(0).getPhysicalNumberOfCells();
		return number;
	}

	public static void destroy() {
		driver.quit();
	}

	/**
	 * Run Step by step of a case on List Test Case in Excel file
	 * 
	 * @param strFileName
	 * @param workbook
	 * @param sheet
	 * @param row
	 * @param desInex
	 * @param expetedInex
	 * @param resultInex
	 * @throws Exception
	 */
	private void stepByStep(String strFileName, int row, int desInex, int expetedInex, int resultInex)
			throws Exception {
		String step = "";
		boolean toDo = true;
		try {
			step = sheet.getRow(row).getCell(desInex).getStringCellValue();
		} catch (NullPointerException e) {
			toDo = false;
		}
		if (toDo) {
			String[] arrStep = step.split("\n");
			String error = "";
			boolean breakResult = false;
			for (String stepCurrent : arrStep) {
				Thread.sleep(100);
				System.err.println(stepCurrent);
				String[] keyWordStep = stepCurrent.split(":");
				String keyWord1 = keyWordStep[0].substring(keyWordStep[0].lastIndexOf(".") + 1, keyWordStep[0].length())
						.trim();
				if (error.length() > 0) {
					cell = sheet.getRow(row).createCell(resultInex);
					cell.setCellValue(error);
					fileOut = new FileOutputStream(strFileName);
					workbook.write(fileOut);
					fileOut.flush();
					fileOut.close();
					System.out.println("ERROR " + error);
					breakResult = true;
					break;
				} else {
					String[] arrBreak = keyWordStep[1].split("\"");
					switch (keyWord1) {
					case "Go to URL":
						String url = stepCurrent.substring(stepCurrent.indexOf("http"), stepCurrent.length() - 1);
						try {
							waitForAjax();
							Thread.sleep(500);
							driver.get(url);
						} catch (Exception e) {
							error = "Pending: wrong " + url;
							e.printStackTrace();
						}
						break;
					case "Navigate to url":
						String urlNavigate = stepCurrent.substring(stepCurrent.indexOf("http"),
								stepCurrent.length() - 1);
						try {
							driver.navigate().to(urlNavigate);
							Thread.sleep(1000);
						} catch (Exception e) {
							error = "Pending: wrong " + urlNavigate;
							e.printStackTrace();
						}
						break;
					case "Click on":
						String elementKey = arrBreak[1];
						try {
							waitForAjax();
							WebElement login = driver.findElement(getIdentifier(elementKey));
							JavascriptExecutor executor = (JavascriptExecutor) driver;
							executor.executeScript("arguments[0].click();", login);
						} catch (Exception e) {
							error = "Pending: wrong " + elementKey + " element";
							e.printStackTrace();
						}
						Thread.sleep(500);
						break;
					case "Scroll Down":
						String elementKeyScroll = arrBreak[1];
						try {
							waitForWebState();
							WebElement scroll = driver.findElement(getIdentifier(elementKeyScroll));
							JavascriptExecutor executor = (JavascriptExecutor) driver;
							executor.executeScript("arguments[0].scrollIntoView(true);", scroll);
						} catch (Exception e) {
							error = "Pending: wrong " + elementKeyScroll + " element";
							e.printStackTrace();
						}
						Thread.sleep(500);
						break;
					case "Input into":
						String elementPlace = arrBreak[1];
						try {
							driver.findElement(getIdentifier(elementPlace)).click();
							String elementInput = arrBreak[3];
							if (elementInput.equals("blank")) {
								driver.findElement(getIdentifier(elementPlace)).clear();
								driver.findElement(getIdentifier(elementPlace)).sendKeys("");
								waitForAjax();
							} else {
								driver.findElement(getIdentifier(elementPlace)).clear();
								driver.findElement(getIdentifier(elementPlace)).sendKeys(elementInput);
								waitForAjax();
							}
						} catch (Exception e) {
							error = "Pending: wrong " + elementPlace + " element";
							e.printStackTrace();
						}
						break;
					case "Input into new value":
						String elementPlace_1 = arrBreak[1];
						try {
							driver.findElement(getIdentifier(elementPlace_1)).click();
							String elementInput_1 = arrBreak[3];
							String dayCurrent = String.valueOf(getCurrentDay());
							driver.findElement(getIdentifier(elementPlace_1)).clear();
							driver.findElement(getIdentifier(elementPlace_1))
									.sendKeys(elementInput_1 + " " + dayCurrent);
							waitForAjax();
						} catch (Exception e) {
							error = "Pending: wrong " + elementPlace_1 + " element";
							e.printStackTrace();
						}
						break;
					case "Wait":
						if (keyWordStep[1].contains("wait for page load successful")) {
							waitForWebState();
						} else {
							String elementTime = arrBreak[1] + "000";
							int result = Integer.parseInt(elementTime);
							try {
								Thread.sleep(result);
							} catch (Exception e) {
								error = "Pending: wrong " + result + " element";
								e.printStackTrace();
							}
						}
						break;
					case "Select value":
						String elementvalue = arrBreak[1];
						String elementvalue1 = arrBreak[3];
						String expected = arrBreak[5];
						try {
							driver.findElement(getIdentifier(elementvalue)).click();
							waitForAjax();
							List<WebElement> lst = driver.findElements(getIdentifier(elementvalue1));
							for (WebElement webElement : lst) {
								String acutalResult = webElement.getText();
								if (acutalResult.equals(expected)) {
									webElement.click();
									break;
								}
							}
						} catch (Exception e) {
							error = "Pending: wrong " + elementvalue + " element";
							e.printStackTrace();
						}
						break;
					case "Wait for":
						driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
						waitForWebState();
						System.out.println(arrBreak[1]);
						try {
							driver.findElements(getIdentifier(arrBreak[1]));
						} catch (Exception e) {
							sheet.getRow(row).getCell(resultInex)
									.setCellValue("Pending: wrong " + arrBreak[1] + " element");
							e.printStackTrace();
						}
						break;
					case "Set checkbox":
						String elementcb = arrBreak[1];
						try {
							driver.findElement(getIdentifier(elementcb)).click();
						} catch (Exception e) {
							error = "Pending: wrong " + elementcb + " element";
							e.printStackTrace();
						}
						break;
					case "Verify check-box":
						String elementUncheck = arrBreak[1];
						try {
							WebElement checkBox = driver.findElement(getIdentifier(elementUncheck));
							if (!checkBox.isSelected()) {
								checkBox.click();
								break;
							} else if (checkBox.isSelected()) {
								break;
							} else {
								error = "Check box invisible";
							}
						} catch (Exception e) {
							error = "Pending: wrong " + elementUncheck + " element";
							e.printStackTrace();
						}
						break;
					case "Select select box":
						try {
							String elementSLB = arrBreak[1];
							String valueSLB = arrBreak[3];
							driver.findElement(getIdentifier(elementSLB)).click();
							List<WebElement> lst = driver.findElements(By.xpath(".//input[@type='search']"));
							WebElement ele = lst.get(lst.size() - 1);
							ele.sendKeys(valueSLB);
							ele.sendKeys(Keys.ENTER);
						} catch (Exception e) {
							error = "Select Box " + arrBreak[1] + " do not work";
							e.printStackTrace();
						}
						break;
					case "Choice text box":
						try {
							String elementTB = arrBreak[1];
							String value = arrBreak[3];
							driver.findElement(getIdentifier(elementTB)).click();
							WebElement ele = driver.findElement(getIdentifier(elementTB));
							ele.sendKeys(value);
							ele.sendKeys(Keys.ENTER);

						} catch (Exception e) {
							e.printStackTrace();
							error = "Text Box " + arrBreak[1] + " do not work";
						}
						break;
					case "Upload image":
						String elementIMG = arrBreak[1];
						try {
							driver.findElement(getIdentifier(elementIMG)).click();
							String elementInput = arrBreak[1];
							String path = arrBreak[3];
							uploadImage(elementInput, path);
						} catch (Exception e) {
							error = "Pending: wrong " + elementIMG + " element";
						}
						break;
					case "Find on":
						String elementFind = arrBreak[1];
						String acutualString = arrBreak[3];
						try {
							List<WebElement> lst = driver.findElements(getIdentifier(elementFind));
							for (WebElement webElement : lst) {
								String acutalResult = webElement.getText();
								System.out.println("Actual String " + acutalResult);
								if (!acutalResult.isEmpty()) {
									if (acutalResult.equals(acutualString)) {
										webElement.click();
										break;
									}
								}
							}
						} catch (Exception e) {
							e.printStackTrace();
							error = "Pending: wrong " + elementFind + " element";
						}
						Thread.sleep(500);
						break;
					case "Remove":
						String elementRemove = arrBreak[1];
						String valueRemove = arrBreak[3];
						boolean checkRemove = false;
						try {
							WebDriverWait wait = new WebDriverWait(driver, 20);
							WebElement elementSelected = wait.until(
									ExpectedConditions.visibilityOf(driver.findElement(getIdentifier(elementRemove))));
							Select listBox = new Select(elementSelected);
							if (listBox.isMultiple()) {
								listBox.deselectByVisibleText(valueRemove);
								Thread.sleep(1000);
								checkRemove = true;
								break;
							} else {
								error = "This slected none multiple" + elementRemove;
							}
							if (!checkRemove) {
								error = "Can not find element " + valueRemove + " on list element " + elementRemove;
							}
						} catch (Exception e) {
							e.printStackTrace();
							error = "Pending: wrong " + elementRemove + " element";
						}
						break;
					case "Choose option":
						try {
							List<WebElement> lst = driver.findElements(getIdentifier(arrBreak[1]));
							for (WebElement e : lst) {
								String value = e.getText().trim();
								System.err.println("Acctual String " + value);
								if (value.equals(arrBreak[3])) {
									waitForAjax();
									e.click();
									break;
								}
							}
						} catch (Exception e) {
							e.printStackTrace();
							error = "Pending: wrong " + arrBreak[1] + " element";
						}
						break;
					case "Click on drop-down":
						try {
							Thread.sleep(300);
							WebElement firstOptions = driver.findElement(getIdentifier(arrBreak[1]));
							firstOptions.click();
							firstOptions.sendKeys(Keys.ENTER);
						} catch (Exception e) {
							e.printStackTrace();
							error = "Pending: wrong " + arrBreak[1] + " element";
						}
						break;
					default:
						error = "Step " + stepCurrent + " has not been defined";
						break;
					}
				}
			}
			if (!breakResult) {
				boolean isResultNull = true;
				try {
					String result = sheet.getRow(row).getCell(expetedInex).getStringCellValue();
					if (result.length() == 0) {
						isResultNull = true;
					} else {
						isResultNull = false;
					}
				} catch (NullPointerException e) {
					isResultNull = true;
				}
				if (isResultNull == false) {
					isResultNull = true;
					String status = "";
					String getResult = sheet.getRow(row).getCell(expetedInex).getStringCellValue();
					String[] result = getResult.split("\n");
					for (String y : result) {
						System.out.println(y);
						String[] keyWordResult = y.split(":");
						String keyWord2 = keyWordResult[0]
								.substring(keyWordResult[0].lastIndexOf(".") + 1, keyWordResult[0].length()).trim();
						String[] arrResult = keyWordResult[1].split("\"");
						switch (keyWord2) {
						case "Should see result":
							String elementFind = arrResult[1];
							String acutualString = arrResult[3];
							try {
								boolean check = false;
								List<WebElement> lst = driver.findElements(getIdentifier(elementFind));
								for (WebElement webElement : lst) {
									String acutalResult = webElement.getText();
									System.out.println("Actual String " + acutalResult);
									if (!acutalResult.isEmpty()) {
										if (acutalResult.equals(acutualString)) {
											status = "PASSED";
											check =true;
											break;
										}
									}
								}
								if (!check) {
									status = "FAILED";
									break;
								}
							} catch (Exception e) {
								e.printStackTrace();
								error = "Pending: wrong " + elementFind + " element";
							}
							Thread.sleep(500);
							break;
						case "Should see message invalid":
							waitForWebState();
							try {
								WebElement elementMessage = driver.findElement(getIdentifier(arrResult[1]));
								String text = elementMessage.getText().trim();
								if (text.contains(arrResult[3])) {
									status = "PASSED";
									break;
								} else {
									status = "FAILED";
								}

							} catch (Exception e) {
								status = "PENDING";
							}
							break;
						case "Should see title is":
							waitForWebState();
							String acctualTitle = driver.getTitle();
							if (acctualTitle.equals(arrResult[1])) {
								status = "PASSED";
								break;
							} else {
								status = "FAILED";
							}
							break;
						default:
							break;
						}

					}

					cell = sheet.getRow(row).createCell(resultInex);
					setColorCell(status, workbook, cell);
					cell.setCellValue(status);
					fileOut = new FileOutputStream(strFileName);
					workbook.write(fileOut);
					fileOut.flush();
					fileOut.close();
				} else {
					XSSFCell cell = null;
					cell = sheet.getRow(row).createCell(resultInex);
					setColorCell("NONE", workbook, cell);
					cell.setCellValue("NONE");
					fileOut = new FileOutputStream(strFileName);
					workbook.write(fileOut);
					fileOut.flush();
					fileOut.close();
				}
			} else {
				breakResult = false;
			}
		}
	}

	/**
	 * Set cell color
	 * 
	 * @param status
	 * @param workbook
	 * @param cell
	 */
	@SuppressWarnings("deprecation")
	public void setColorCell(String status, XSSFWorkbook workbook, XSSFCell cell) {
		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(HSSFColor.BLACK.index);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(HSSFColor.BLACK.index);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(HSSFColor.BLACK.index);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(HSSFColor.BLACK.index);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
		switch (status) {
		case "PASSED":
			font.setColor(HSSFColor.WHITE.index);
			style.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			style.setFont(font);
			cell.setCellStyle(style);
			break;
		case "FAILED":
			font.setColor(HSSFColor.WHITE.index);
			style.setFillForegroundColor(HSSFColor.RED.index);
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			style.setFont(font);
			cell.setCellStyle(style);
			break;
		case "PENDING":
			font.setColor(HSSFColor.WHITE.index);
			style.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			style.setFont(font);
			cell.setCellStyle(style);
			break;
		default:
			font.setColor(HSSFColor.BLACK.index);
			style.setFont(font);
			cell.setCellStyle(style);
			break;
		}
	}

	public static void uploadImage(String imageName, String path) throws AWTException {
		Robot robot = new Robot();
		path = path.replaceAll("\\", "\\\\");
		String pathFileName = path + imageName;
		StringSelection stringselection = new StringSelection(pathFileName);
		Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringselection, null);
		robot.setAutoDelay(1000);
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_V);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_V);
		robot.setAutoDelay(1000);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);

	}

	public void tearDown() {
		driver.quit();
	}

	public void captureFullScreen(int numberOfTC) throws IOException, InterruptedException {
		Screenshot screenshot = new AShot().shootingStrategy(ShootingStrategies.viewportPasting(10000))
				.takeScreenshot(driver);
		ImageIO.write(screenshot.getImage(), "PNG", new File(data.getStoreScreenShoot(numberOfTC + ".png")));

	}

	private void CheckLanguage(String strFileName, XSSFWorkbook workbook, XSSFSheet sheet, int row, int desInex,
			int expetedInex, int resultInex) throws Exception {
		String step = "";
		boolean toDo = true;
		try {
			step = sheet.getRow(row).getCell(desInex).getStringCellValue();
		} catch (NullPointerException e) {
			toDo = false;
		}
		if (toDo) {
			String[] arrStep = step.split("\n");
			for (String stepCurrent : arrStep) {
				Thread.sleep(500);
				System.err.println(stepCurrent);
				String[] keyWordStep = stepCurrent.split(":");
				String keyWord1 = keyWordStep[0].substring(keyWordStep[0].lastIndexOf(".") + 1, keyWordStep[0].length())
						.trim();
				String[] arrBreak = null;
				switch (keyWord1) {
				case "Go to URL":
					String url = stepCurrent.substring(stepCurrent.indexOf("http"), stepCurrent.length() - 1);
					try {
						Thread.sleep(3000);
						driver.get(url);
						Thread.sleep(1000);
					} catch (Exception e) {
						e.printStackTrace();
					}
					break;
				case "Navigate to url":
					String urlNavigate = stepCurrent.substring(stepCurrent.indexOf("http"), stepCurrent.length() - 1);
					try {
						driver.navigate().to(urlNavigate);
						Thread.sleep(1000);
					} catch (Exception e) {
						e.printStackTrace();
					}
					break;
				case "Click on":
					arrBreak = keyWordStep[1].split("\"");
					String elementClickOn = arrBreak[1];
					try {
						waitForAjax();
						WebElement login = driver.findElement(getIdentifier(elementClickOn));
						JavascriptExecutor executor = (JavascriptExecutor) driver;
						executor.executeScript("arguments[0].click();", login);
					} catch (Exception e) {
						e.printStackTrace();
					}
					Thread.sleep(500);
					break;
				case "Scroll Down":
					arrBreak = keyWordStep[1].split("\"");
					String elementKeyScroll = arrBreak[1];
					try {
						waitForAjax();
						WebElement scroll = driver.findElement(getIdentifier(elementKeyScroll));
						JavascriptExecutor executor = (JavascriptExecutor) driver;
						executor.executeScript("arguments[0].scrollIntoView(true);", scroll);
					} catch (Exception e) {
						e.printStackTrace();
					}
					Thread.sleep(500);
					break;
				case "Input into":
					arrBreak = keyWordStep[1].split("\"");
					String elementInputInto = arrBreak[1];
					try {
						waitForAjax();
						driver.findElement(getIdentifier(elementInputInto)).click();
						String elementInput = arrBreak[3];
						if (elementInput.equals("blank")) {
							driver.findElement(getIdentifier(elementInputInto)).clear();
							driver.findElement(getIdentifier(elementInputInto)).sendKeys("");
						} else {
							driver.findElement(getIdentifier(elementInputInto)).clear();
							driver.findElement(getIdentifier(elementInputInto)).sendKeys(arrBreak[3]);
						}
					} catch (Exception e) {
						e.printStackTrace();
					}
					break;
				case "Wait":
					if (keyWordStep[1].contains("wait for page load successful")) {
						waitForAjax();
						waitForWebState();
					} else {
						arrBreak = keyWordStep[1].split("\"");
						String elementTime = arrBreak[1] + "000";
						int result = Integer.parseInt(elementTime);
						try {
							Thread.sleep(result);
						} catch (Exception e) {
							e.printStackTrace();
						}
					}
					break;
				case "Wait for":
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					arrBreak = keyWordStep[1].split("\"");
					waitForWebState();
					try {
						driver.findElements(getIdentifier(arrBreak[1]));
					} catch (Exception e) {
						sheet.getRow(row).getCell(resultInex)
								.setCellValue("Pending: wrong " + arrBreak[1] + " element");
						e.printStackTrace();
					}
					break;
				case "Verify check-box":
					arrBreak = keyWordStep[1].split("\"");
					String elementUncheck = arrBreak[1];
					try {
						WebElement checkBox = driver.findElement(getIdentifier(elementUncheck));
						if (!checkBox.isSelected()) {
							checkBox.click();
							break;
						} else if (checkBox.isSelected()) {
							break;
						} else {
						}
					} catch (Exception e) {
						e.printStackTrace();
					}
					break;
				case "Find on":
					arrBreak = keyWordStep[1].split("\"");
					String elementFind = arrBreak[1];
					String acutualString = arrBreak[3];
					try {
						List<WebElement> lst = driver.findElements(getIdentifier(elementFind));
						for (WebElement webElement : lst) {
							String acutalResult = webElement.getText();
							if (!acutalResult.isEmpty()) {
								if (acutalResult.equals(acutualString)) {
									webElement.click();
									break;
								}
							}
						}

					} catch (Exception e) {
						e.printStackTrace();
					}
					Thread.sleep(1000);
					break;
				}
			}
		}
	}

	// Get Current daytime
	public LocalDateTime getCurrentDay() {
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
		LocalDateTime now = LocalDateTime.now();
		System.out.println(dtf.format(now));
		return now;
	}

	public void CheckLanguageLookup(String fileName, String sheetName, String numberRound, DataTable data)
			throws Exception {
		List<List<String>> list = data.raw();
		String rowDo = list.get(0).get(0).trim();
		DataProvider dataPro = new DataProvider();
		// Read file excel
		String strFileName = dataPro.getExcelfile(fileName);
		FileInputStream file = new FileInputStream(strFileName);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		int desInex = 0, expetedInex = 0, resultInex = 0;
		int columnCount = this.getCellCount(strFileName, sheetName);
		int roundResult = Integer.parseInt(numberRound);
		for (int col = 0; col < columnCount; col++) {
			String titleColumnText = sheet.getRow(0).getCell(col).toString().trim();
			if (titleColumnText.equalsIgnoreCase("Description")) {
				desInex = col;
			}
			if (titleColumnText.equalsIgnoreCase("Expected Result")) {
				expetedInex = col;
			}
			if (titleColumnText.equalsIgnoreCase("Result Round " + roundResult)) {
				resultInex = col;
			}
		}
		if (rowDo.equals("All")) {
			int rowCount = sheet.getLastRowNum();
			for (int i = 1; i < rowCount; i++) {
				CheckLanguage(strFileName, workbook, sheet, i, desInex, expetedInex, resultInex);
			}
		} else {
			int rowCount = sheet.getLastRowNum();
			if (rowDo.contains(",")) { // Case 2: Run some row on Excel file
				String[] arrRowDo = rowDo.split(",");
				for (int i = 0; i < arrRowDo.length; i++) {
					int row = Integer.parseInt(arrRowDo[i]);
					if (row > rowCount) {
						Assert.assertTrue("Index of row do out of range ", false);
					} else {
						CheckLanguage(strFileName, workbook, sheet, row, desInex, expetedInex, resultInex);
					}

				}
			} else { // Case 3: Run a row on Excel file
				int row = Integer.parseInt(rowDo.trim());
				if (row > rowCount) {
					Assert.assertTrue("Index of row do out of range ", false);
				} else {
					CheckLanguage(strFileName, workbook, sheet, row, desInex, expetedInex, resultInex);
				}

			}
		}
	}

	public int getResponse(String urlString) throws MalformedURLException, IOException {
		int reponseCode = 0;
		try {
			URL url = new URL(urlString);
			HttpURLConnection huc = (HttpURLConnection) url.openConnection();
			if (!ReadProperties.getElementsAndUrls("auuser").equals("")
					&& !ReadProperties.getElementsAndUrls("aupass").equals("")) {
				String author = ReadProperties.getElementsAndUrls("auuser") + ":"
						+ ReadProperties.getElementsAndUrls("aupass");
				String encoding = Base64.encodeBytes(author.getBytes());
				huc.setRequestProperty("Authorization", "Basic " + encoding);
			}
			huc.setRequestMethod("HEAD");
			huc.connect();
			reponseCode = huc.getResponseCode();
		} catch (Exception e) {
			e.getMessage();
		}
		return reponseCode;
	}

	public void addImageEvidence() {

	}

	/**
	 * Check time load page Read url frome excel file public void
	 * checkTimePageLoad(String sheetName, String excelName) throws IOException,
	 * AWTException, InterruptedException { int response = getResponse("url"); if
	 * (!url.equals("")) { if (response == 200) { openPageToCheckPageSpeed(url,
	 * myScenario); Long loadtime = (Long) ((JavascriptExecutor)
	 * driver).executeScript( return performance.timing.loadEventEnd -
	 * performance.timing.navigationStart;); System.out.println(loadtime + "(ms)");
	 * setValueInExcel(i, 1, loadtime + "(ms)", ""); } else { setValueInExcel(i, 1,
	 * "The link is broken. Status is " + response, "YELLOW"); } } }
	 **/

}
