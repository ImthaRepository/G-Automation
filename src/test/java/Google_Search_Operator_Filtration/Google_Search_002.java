package Google_Search_Operator_Filtration;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import org.apache.log4j.Logger;
import org.apache.log4j.PatternLayout;
import org.apache.log4j.RollingFileAppender;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;


public class Google_Search_002 {
	public String getPropertyFileValue(String key) throws FileNotFoundException, IOException {
		Properties properties = new Properties();
		properties.load(new FileInputStream(System.getProperty("user.dir") + "\\input.properties"));
		Object object = properties.get(key);
		String value = (String) object;
		return value;

	}

	public static String stamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss").format(new Date());

	WebDriver driver;

	@BeforeTest
	public void browseropen() throws FileNotFoundException, IOException {
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		String url=getPropertyFileValue("url");
		driver.get(url);
		logger.info("URL Entered");
		System.out.println("URL Entered");

	}

	@Test(priority = 4, dependsOnMethods = "analyseURLs")
	public void signOut() throws InterruptedException {
		driver.findElement(By.xpath("//img[contains(@class,'global-nav__me-photo')]")).click();
		System.out.println("Profile Icon clicked");
		logger.info("Profile Icon Clicked");
		Thread.sleep(1500);
		WebElement signOutBtn = driver.findElement(By.xpath("//a[text()= 'Sign Out']"));
		JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].click();", signOutBtn);
		System.out.println("Sign out button clicked");
		logger.info("Sign out button clicked");
		try {
			driver.findElement(By.xpath("//button//span[text()= 'Sign out']")).click();
			System.out.println("Sign out in the pop up is clicked");
			logger.info("Sign out in the pop up is clicked");
		} catch (Exception e) {
			System.out.println("Sign out pop up not available");
			logger.info("Sign out pop up not available");
		}
		
		
	}
	
	//@AfterTest
	public void closebrowser() {

		 if (driver != null) {
	            try {
	                driver.quit(); // Properly close session
	            } catch (WebDriverException e) {
	                System.err.println("WebDriverException caught during quit(): " + e.getMessage());
	                // You can log or handle Connection Reset specifically
	                if (e.getMessage().contains("Connection reset")) {
	                    // handle or log specifically
	                }
	            } catch (Exception e) {
	                System.err.println("Unexpected error during driver quit: " + e.getMessage());
	            }
	        }
	}

	@Test(priority = 1)
	public void Fetchdata() throws InterruptedException, FileNotFoundException, IOException {
		// driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));

		
		WebDriverWait waiting = new WebDriverWait(driver, Duration.ofSeconds(40));

		WebElement SearchField = driver.findElement(By.xpath("//textarea[@aria-label='Search']"));
		SearchField.sendKeys(getPropertyFileValue("Query"), Keys.ENTER);
		System.out.println("Query Entered");
		try {
			waiting.until(ExpectedConditions
					.visibilityOfElementLocated(By.xpath("//div//h1[text()='Search Results']/parent::div//a")));
			createExcelFile("Company Raw files", stamp);
			try {
				int pageElement = driver.findElements(By.xpath("//h1[text()='Page navigation']/parent::div//td//a")).size();
				
				for (int j = 1; j < pageElement; j++) {
					int searchResultElement = driver.findElements(By.xpath("//div//h1[text()='Search Results']/parent::div//a"))
							.size();
					System.out.println("Total Link Count is - " + searchResultElement);
					logger.info("Total Link Count is - " + searchResultElement);
					String filepath = System.getProperty("user.dir") + "\\Company Raw files\\" + stamp + ".xlsx";
					for (int i = 1; i <= searchResultElement; i++) {
						WebElement searchLinksElement = driver
								.findElement(By.xpath("(//div//h1[text()='Search Results']/parent::div//a)[" + i + "]"));
						String searchlinks = searchLinksElement.getDomAttribute("href");
					      
					    	  writeInMasterSheet("Links", 0, searchlinks, filepath, "Sheet1");
					        } 
						Thread.sleep(1500);
					}
					
				
				deleteExcelIfNoSheets();
				createExcelFile("Company Filtered files",stamp+"-Company Details");
			} catch (Exception e) {
				int searchResultElement = driver.findElements(By.xpath("//div//h1[text()='Search Results']/parent::div//a"))
						.size();
				System.out.println("Total link count is - " + searchResultElement);
				logger.info("Total link count is - " + searchResultElement);
				String filepath = System.getProperty("user.dir") + "\\Company Raw files\\" + stamp + ".xlsx";
				for (int i = 1; i <= searchResultElement; i++) {
					WebElement searchLinksElement = driver
							.findElement(By.xpath("(//div//h1[text()='Search Results']/parent::div//a)[" + i + "]"));
					String searchlinks = searchLinksElement.getDomAttribute("href");
				
				    	  writeInMasterSheet("Links", 0, searchlinks, filepath, "Sheet1");
				         
					
				}
				deleteExcelIfNoSheets();
				createExcelFile("Company Filtered files",stamp+"-Company Details");
			}
			
			
		} catch (Exception e) {
			System.out.println("NO Search Result Found for this filter");
			logger.info("NO Search Result Found for this filter");
		}

	}

	// -------------------------------------------------Fetched link Validation and Filter----------------------------------------------
	static int iterateFounder = 1;
	static int iterateCEO=1;
	static int iterateCTO=1;
	static int newRow=0;
	@Test(priority = 2, dataProvider = "excelData", dependsOnMethods = "Fetchdata")
	public void analyseURLs(String URL) throws IOException {
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		WebDriverWait waiting = new WebDriverWait(driver, Duration.ofSeconds(5));

		String filepath=System.getProperty("user.dir") + "\\Company Filtered files\\"+stamp+"-Company Details.xlsx";
		if (URL.contains("linkedin.com")) {
		driver.get(URL);
		try {
//			try {
//				Thread.sleep(1000);
//				//wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//icon[contains(@class,'modal__modal-dismiss-icon')])[1]")));
//				driver.findElement(By.xpath("(//icon[contains(@class,'modal__modal-dismiss-icon')])[1]")).click();
//			} catch (Exception e) {
//				driver.findElement(By.xpath("(//icon[contains(@class,'modal__modal-dismiss-icon')])[2]")).click();
//			} finally {
//				System.out.println("Sign In Pop is not available");
//			}
//			try {
//	
//				//wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//icon[contains(@class,'modal__modal-dismiss-icon')])[1]")));
//				driver.findElement(By.xpath("(//icon[contains(@class,'modal__modal-dismiss-icon')])[3]")).click();
//			} catch (Exception e) {
//				driver.findElement(By.xpath("(//icon[contains(@class,'modal__modal-dismiss-icon')])[4]")).click();
//			} finally {
//				System.out.println("NO Sign In Pop available");
//			}
//			
			try {
				 Actions actions = new Actions(driver);

			        // Move to (50px, 100px) from the top-left of the page and click
			        actions.moveByOffset(50, 100).click().perform();

			        // Optional: Reset the mouse pointer to avoid affecting future actions
			        actions.moveByOffset(-50, -100).perform(); // moves back
			} catch (Exception e) {
				System.out.println("No Sign in popup available");			}
			driver.findElement(
					By.xpath("//a[contains(text(), 'Join now')]/following-sibling::a[contains(text(), 'Sign in')]"))
					.click();

			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("username")));
			WebElement usernameField = driver.findElement(By.id("username"));
			usernameField.clear();
			usernameField.sendKeys(getPropertyFileValue("username"));
			System.out.println("username Entered");
			logger.info("Username Entered");

			WebElement passwordField = driver.findElement(By.id("password"));
			passwordField.clear();
			passwordField.sendKeys(getPropertyFileValue("password"));
			System.out.println("Password Entered");
			logger.info("password Entered");

			driver.findElement(By.xpath("//button[@aria-label='Sign in']")).click();
			System.out.println("Sign in Clicked");
			logger.info("Sign In clicked");
		} catch (Exception e) {
			System.out.println("Login Not Needed");
		}
		try {
			wait.until(ExpectedConditions.visibilityOfElementLocated(
					By.xpath("//div[contains(@class,'artdeco-card')]//div[contains(@class,'company-name')]//a")));
			WebElement compName = driver.findElement(
					By.xpath("//div[contains(@class,'artdeco-card')]//div[contains(@class,'company-name')]//a"));
			String companyName=compName.getText();
			String companyURL=compName.getAttribute("href");
			writeCompanyDatas(filepath, "Sheet1", companyName,companyURL);
			compName.click();
				

			try {
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li[contains(@class,'org-page-navigation_')]//a[text()='People']")));
				WebElement peopleElement = driver.findElement(By.xpath("//li[contains(@class,'org-page-navigation_')]//a[text()='People']"));
				peopleElement.click();
				System.out.println("People Element Clicked");
				logger.info("People toggle Clicked");
				
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//textarea[@id='people-search-keywords']")));
				WebElement peopleSearchField = driver.findElement(By.xpath("//textarea[@id='people-search-keywords']"));
				try {
					peopleSearchField.sendKeys("Founder",Keys.ENTER);
					waiting.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[contains(@class,'org-people-profile-card__profile-info')]")));
					int profileCount = driver.findElements(By.xpath("//div[contains(@class,'org-people-profile-card__profile-info')]")).size();
					for (int i = 1; i <= profileCount; i++) {
						WebElement profileDetailsElement = driver.findElement(By.xpath("(//div[contains(@class,'org-people-profile-card__profile-info')]//div[contains(@class,'entity-lockup__content')])["+i+"]"));
						String profileDetail=profileDetailsElement.getText();
						System.out.println("Profile Datails - "+profileDetail);
						logger.info("profileDetail - "+profileDetail);
						try {
							WebElement profileLinkElement = driver.findElement(By.xpath("(//div[contains(@class,'org-people-profile-card__profile-info')]//div[contains(@class,'entity-lockup__content')]//a)["+i+"]"));
							String profileLink=profileLinkElement.getAttribute("href");
							System.out.println("Profile link - "+profileLink);
							logger.info("profile link -"+profileLink);
							writeDataInColumns3and4(filepath, "Sheet1", profileDetail, profileLink, iterateFounder);
							iterateFounder++;
						} catch (Exception e) {
							System.out.println("Profile link is disabled");
							logger.info("Profile link is disabled");
							writeDataInColumns3and4(filepath, "Sheet1", profileDetail, "-", iterateFounder);
							iterateFounder++;
						}
						
					}
					
				} catch (Exception e) {
					writeDataInColumns3and4(filepath, "Sheet1", "-", "-", iterateFounder);
					System.out.println("No Founder Profile found for this company");
					logger.info("No Founder Profile found for this company");			
				}
				
				try {
					
					driver.findElement(By.xpath("//button//span[text()='Founder']")).click();
					
					Thread.sleep(2000);
					WebElement peopleSearchField1 = driver.findElement(By.xpath("//textarea[@id='people-search-keywords']"));
					peopleSearchField1.click();
					System.out.println("CEO clicked");
					peopleSearchField1.sendKeys("CEO",Keys.ENTER);
					System.out.println("CEO Entered");
					waiting.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[contains(@class,'org-people-profile-card__profile-info')]")));
					int profileCount = driver.findElements(By.xpath("//div[contains(@class,'org-people-profile-card__profile-info')]")).size();
					for (int i = 1; i <= profileCount; i++) {
						WebElement profileDetailsElement = driver.findElement(By.xpath("(//div[contains(@class,'org-people-profile-card__profile-info')]//div[contains(@class,'entity-lockup__content')])["+i+"]"));
						String profileDetail=profileDetailsElement.getText();
						System.out.println("Profile Datails - "+profileDetail);
						logger.info("profileDetail - "+profileDetail);
						try {
							WebElement profileLinkElement = driver.findElement(By.xpath("(//div[contains(@class,'org-people-profile-card__profile-info')]//div[contains(@class,'entity-lockup__content')]//a)["+i+"]"));
							String profileLink=profileLinkElement.getAttribute("href");
							System.out.println("Profile link - "+profileLink);
							logger.info("profile link -"+profileLink);
							writeDataInColumns5and6(filepath, "Sheet1", profileDetail, profileLink, iterateCEO);
							iterateCEO++;
						} catch (Exception e) {
							System.out.println("Profile link is disabled");
							logger.info("Profile link is disabled");
							writeDataInColumns5and6(filepath, "Sheet1", profileDetail, "-", iterateCEO);
							iterateCEO++;
						}
						
					}
					
				} catch (Exception e) {
					e.printStackTrace();
					writeDataInColumns5and6(filepath, "Sheet1", "-", "-", iterateCEO);
					System.out.println("The error is" + e.getMessage());
					System.out.println("CEO Profile Not found for this company");
					logger.info("No CEO Profile found for this company");
				}
				
				try {
					
					driver.findElement(By.xpath("//button//span[text()='CEO']")).click();
					Thread.sleep(2000);
					//wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//textarea[@id='people-search-keywords']")));
					WebElement peopleSearchField2 = driver.findElement(By.xpath("//textarea[@id='people-search-keywords']"));
					
					peopleSearchField2.sendKeys("CTO",Keys.ENTER);
					System.out.println("CEO Entered");
					waiting.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[contains(@class,'org-people-profile-card__profile-info')]")));
					int profileCount = driver.findElements(By.xpath("//div[contains(@class,'org-people-profile-card__profile-info')]")).size();
					for (int i = 1; i <= profileCount; i++) {
						WebElement profileDetailsElement = driver.findElement(By.xpath("(//div[contains(@class,'org-people-profile-card__profile-info')]//div[contains(@class,'entity-lockup__content')])["+i+"]"));
						String profileDetail=profileDetailsElement.getText();
						System.out.println("Profile Datails - "+profileDetail);
						logger.info("profileDetail - "+profileDetail);
						try {
							WebElement profileLinkElement = driver.findElement(By.xpath("(//div[contains(@class,'org-people-profile-card__profile-info')]//div[contains(@class,'entity-lockup__content')]//a)["+i+"]"));
							String profileLink=profileLinkElement.getAttribute("href");
							writeDataInColumns7and8(filepath, "Sheet1", profileDetail, profileLink, iterateCTO);
							iterateCTO++;
							System.out.println("Profile link - "+profileLink);
							logger.info("profile link -"+profileLink);
						} catch (Exception e) {
							System.out.println("Profile link is disabled");
							logger.info("Profile link is disabled");
							writeDataInColumns7and8(filepath, "Sheet1", profileDetail, "-", iterateCTO);
							iterateCTO++;
						}
						
					}
				} catch (Exception e2) {
					e2.printStackTrace();
					writeDataInColumns7and8(filepath, "Sheet1", "-", "-", iterateCTO);
					System.out.println("The error is" + e2.getMessage());
					System.out.println("NO CTO Profile found for this company");
					logger.info("CTO Profile not found for this company");
				}
				
				
				
		
			} catch (Exception e) {
			
				System.out.println("No Data Found for this Filter");
				logger.info("No Data Found for this filter");

		}

			newRow = getFirstEmptyRowIndexIn8Columns(filepath, "Sheet1");
			iterateCEO=newRow;
			iterateFounder=newRow;
			iterateCTO=newRow;
			autoFitColumn(filepath, "Sheet1");
		} catch (Exception e) {
			System.out.println("No Company Profile is found");
			logger.info("No Company Profile is found");
		}
		}else {
			
	            System.out.println("The links is not from linkedIn - " + URL);
	            logger.info("The links is not from linkedIn - "+ URL);
	       
		}
	}
	//----------------------------------------------------Delete Excel file if it has no Sheet -----------------------------------------------	
		@Test(priority = 3, dependsOnMethods = "analyseURLs")
		public void deleteExcelIfNoSheet() throws FileNotFoundException, IOException {
		
			String filePath = System.getProperty("user.dir") + "\\Company Filtered files\\" + stamp + "-Company Details.xlsx"; // Replace with actual file path
			File file = new File(filePath);

			if (!file.exists()) {
				System.out.println("File does not exist.");
				return;
			}

			try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {

				int numberOfSheets = workbook.getNumberOfSheets();
				if (numberOfSheets == 0) {
					// Close workbook before deleting file
					workbook.close();

					if (file.delete()) {
						System.out.println("Excel file deleted as it contains no sheets.");
						logger.info("Excel file deleted as it contains no sheets.");
					} else {
						System.out.println("Failed to delete the file.");
						logger.info("Failed to delete the file.");
					}
				} else {
					System.out.println("Excel file has sheets, not deleted.");
					logger.info("Excel file has sheets, not deleted.");
					countFilledCellsInExcel(filePath);
					autoFitColumn(filePath, "Sheet1");
				}

			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		
		public void deleteExcelIfNoSheets() throws FileNotFoundException, IOException {
			String filePath = System.getProperty("user.dir") + "\\Company Raw files\\" +stamp+".xlsx"; // Replace with actual file path
			File file = new File(filePath);
			if (!file.exists()) {
				System.out.println("File does not exist.");
				logger.info("File does not exist.");
				return;
			}

			try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {

				int numberOfSheets = workbook.getNumberOfSheets();
				if (numberOfSheets == 0) {
					// Close workbook before deleting file
					workbook.close();

					if (file.delete()) {
						System.out.println("Excel file deleted as it contains no sheets.");
						logger.info("Excel file deleted as it contains no sheets.");
						Assert.fail();
					} else {
						System.out.println("Failed to delete the file.");
						logger.info("Failed to delete the file.");
					}
				} else {
					System.out.println("Excel file has sheets, not deleted.");
					logger.info("Excel file has sheets, not deleted.");
				}

			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	// ----------------------------------------------Data Provider Code--------------------------------------------------------------------
	@DataProvider(name = "excelData")
	public Object[][] getDataFromExcel() throws FileNotFoundException, IOException {
		String filePath = System.getProperty("user.dir") + "\\Company Raw files\\" + stamp +".xlsx";
		String sheetName = "Sheet1";
		return readExcelcomData(filePath, sheetName);
	}

	
	private Object[][] readExcelcomData(String filePath, String sheetName) {
		List<Object[]> dataList = new ArrayList();

		try (FileInputStream fis = new FileInputStream(new File(filePath));
				Workbook workbook = WorkbookFactory.create(fis)) {

			Sheet sheet = workbook.getSheet(sheetName);
			Iterator<Row> rows = sheet.iterator();

			while (rows.hasNext()) {
				Row row = rows.next();
				List<String> cellList = new ArrayList<>();

				boolean skipRow = false;

				for (Cell cell : row) {
					cell.setCellType(CellType.STRING);
					String value = cell.getStringCellValue().trim();

					if (value.isEmpty()) {
						skipRow = true; // Skip if any cell in the row is empty
						break;
					}

					cellList.add(value);
				}

				if (!skipRow && !cellList.isEmpty()) {
					dataList.add(cellList.toArray(new Object[0]));
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

		return dataList.toArray(new Object[0][]);
	}
		
	

	
//------------------------------------------------------------ Insert Value in the Excel Sheet --------------------------------------------//

	


	public static void writeInMasterSheet(String sheetName, int columnNumber, String data, String fileName,
			String sheetname) throws IOException {
		File file = new File(System.getProperty("user.dir") + "//Master Sheet.xlsx");
		Workbook workbook;
		Sheet sheet;

		// Load existing workbook or create new one
		if (file.exists()) {
			FileInputStream fileInputStream = new FileInputStream(file);
			workbook = new XSSFWorkbook(fileInputStream);
			fileInputStream.close();
		} else {
			workbook = new XSSFWorkbook();
		}

		// Get or create sheet
		sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			sheet = workbook.createSheet(sheetName);
		}

		// Check for duplicate in the target column
		boolean isDuplicate = false;
		int lastRowNum = sheet.getLastRowNum();
		for (int i = 0; i <= lastRowNum; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				Cell cell = row.getCell(columnNumber);
				if (cell != null && cell.getCellType() == CellType.STRING) {
					if (cell.getStringCellValue().equalsIgnoreCase(data)) {
						isDuplicate = true;
						System.out.println("Duplicate entry in the Mastersheet");
						logger.info("Duplicate entry in the Mastersheet");
						break;
					}
				}
			}
		}

		// Write the data if not duplicate
		if (!isDuplicate) {
			Row newRow = sheet.createRow(lastRowNum + 1);
			Cell newCell = newRow.createCell(columnNumber);
			newCell.setCellValue(data);
			System.out.println("Data Written in Mastersheet - " + data);
			logger.info("Data Written in Mastersheet - " + data);
			writeData(fileName, sheetname, data);

		}

		// Write to file
		FileOutputStream fileOutputStream = new FileOutputStream(file);
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		workbook.close();
	}
	
	public static void writeCompanyDatas(String fileName, String sheetName, String dataCol1, String dataCol2) {
	    String filePath = fileName;

	    // Static titles for column 1 and 2
	    String[] columnTitles = {
	        "Company Name", // Column 1 (index 0)
	        "Company URL"   // Column 2 (index 1)
	    };

	    try (FileInputStream fis = new FileInputStream(filePath);
	         Workbook workbook = new XSSFWorkbook(fis)) {

	        Sheet sheet = workbook.getSheet(sheetName);
	        if (sheet == null) {
	            sheet = workbook.createSheet(sheetName);
	        }

	        // Create or update the header row at index 0
	        Row headerRow = sheet.getRow(0);
	        if (headerRow == null) {
	            headerRow = sheet.createRow(0);
	        }

	        // Write titles for column 1 and 2
	        for (int i = 0; i < columnTitles.length; i++) {
	            Cell cell = headerRow.getCell(i);
	            if (cell == null || cell.getCellType() == CellType.BLANK) {
	                cell = headerRow.createCell(i);
	                cell.setCellValue(columnTitles[i]);
	            }
	        }

	        // Find first completely empty row (all first 8 columns blank)
	        int lastRow = sheet.getLastRowNum();
	        int writeRowIndex = -1;

	        for (int i = 1; i <= lastRow + 1; i++) { // Start from 1 to skip header
	            Row row = sheet.getRow(i);
	            if (row == null) {
	                writeRowIndex = i;
	                break;
	            }

	            boolean isEmpty = true;
	            for (int j = 0; j < 8; j++) {
	                Cell cell = row.getCell(j);
	                if (cell != null && cell.getCellType() != CellType.BLANK) {
	                    isEmpty = false;
	                    break;
	                }
	            }

	            if (isEmpty) {
	                writeRowIndex = i;
	                break;
	            }
	        }

	        // Write data to the found empty row
	        if (writeRowIndex != -1) {
	            Row dataRow = sheet.getRow(writeRowIndex);
	            if (dataRow == null) {
	                dataRow = sheet.createRow(writeRowIndex);
	            }

	            dataRow.createCell(0).setCellValue(dataCol1); // Column 1
	            dataRow.createCell(1).setCellValue(dataCol2); // Column 2

	            System.out.println("Data written at row index: " + writeRowIndex + ", columns 1 & 2");
	            logger.info("Data written at row index: " + writeRowIndex + ", columns 1 & 2");
	        } else {
	            System.out.println("No empty row found to write data.");
	            logger.warn("No empty row found to write data.");
	        }

	        // Save the updated Excel file
	        try (FileOutputStream fos = new FileOutputStream(filePath)) {
	            workbook.write(fos);
	        }

	    } catch (IOException e) {
	        e.printStackTrace();
	    }
	}

	   public static void writeDataInColumns3and4(String filePath, String sheetName, String dataCol3, String dataCol4, int iteration) {
	        try (FileInputStream fis = new FileInputStream(filePath);
	             Workbook workbook = new XSSFWorkbook(fis)) {

	            Sheet sheet = workbook.getSheet(sheetName);
	            if (sheet == null) {
	                sheet = workbook.createSheet(sheetName);
	            }

	            // Create header row if it doesn't exist
	            Row headerRow = sheet.getRow(0);
	            if (headerRow == null) {
	                headerRow = sheet.createRow(0);
	            }

	            // Set headers in columns 3 and 4
	            Cell headerCellCol3 = headerRow.getCell(2); // Column C
	            if (headerCellCol3 == null || headerCellCol3.getCellType() == CellType.BLANK) {
	                headerCellCol3 = headerRow.createCell(2);
	                headerCellCol3.setCellValue("Founder Details");
	            }

	            Cell headerCellCol4 = headerRow.getCell(3); // Column D
	            if (headerCellCol4 == null || headerCellCol4.getCellType() == CellType.BLANK) {
	                headerCellCol4 = headerRow.createCell(3);
	                headerCellCol4.setCellValue("Founder links");
	            }

	            // Create target row if it doesn't exist
	            Row row = sheet.getRow(iteration);
	            if (row == null) {
	                row = sheet.createRow(iteration);
	            }

	            // Write data to columns 3 and 4
	            Cell cellCol3 = row.createCell(2); // Column C
	            cellCol3.setCellValue(dataCol3);

	            Cell cellCol4 = row.createCell(3); // Column D
	            cellCol4.setCellValue(dataCol4);

	            // Close input stream before writing back
	            fis.close();

	            // Save the changes to file
	            try (FileOutputStream fos = new FileOutputStream(filePath)) {
	                workbook.write(fos);
	            }

	            workbook.close();
	            System.out.println("Data written successfully to row " + iteration + " in columns 3 and 4.");

	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }

	    public static void writeDataInColumns5and6(String filePath, String sheetName, String dataCol5, String dataCol6, int iteration) {
	        try (FileInputStream fis = new FileInputStream(filePath);
	             Workbook workbook = new XSSFWorkbook(fis)) {

	            Sheet sheet = workbook.getSheet(sheetName);
	            if (sheet == null) {
	                sheet = workbook.createSheet(sheetName);
	            }

	            // Create or get header row
	            Row headerRow = sheet.getRow(0);
	            if (headerRow == null) {
	                headerRow = sheet.createRow(0);
	            }

	            // Set header titles for column 5 and 6
	            Cell headerCellCol5 = headerRow.getCell(4); // Column E
	            if (headerCellCol5 == null || headerCellCol5.getCellType() == CellType.BLANK) {
	                headerCellCol5 = headerRow.createCell(4);
	                headerCellCol5.setCellValue("CEO Details");
	            }

	            Cell headerCellCol6 = headerRow.getCell(5); // Column F
	            if (headerCellCol6 == null || headerCellCol6.getCellType() == CellType.BLANK) {
	                headerCellCol6 = headerRow.createCell(5);
	                headerCellCol6.setCellValue("CEO Link");
	            }

	            // Create or get target row based on iteration
	            Row row = sheet.getRow(iteration);
	            if (row == null) {
	                row = sheet.createRow(iteration);
	            }

	            // Write data to column 5 and 6
	            Cell cellCol5 = row.createCell(4); // Column E
	            cellCol5.setCellValue(dataCol5);

	            Cell cellCol6 = row.createCell(5); // Column F
	            cellCol6.setCellValue(dataCol6);

	            fis.close(); // Close input stream before writing

	            // Save changes to file
	            try (FileOutputStream fos = new FileOutputStream(filePath)) {
	                workbook.write(fos);
	            }

	            workbook.close();
	            System.out.println("Data written to row " + iteration + " in columns 5 and 6.");

	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	    public static void writeDataInColumns7and8(String filePath, String sheetName, String dataCol7, String dataCol8, int iteration) {
	        try (FileInputStream fis = new FileInputStream(filePath);
	             Workbook workbook = new XSSFWorkbook(fis)) {

	            Sheet sheet = workbook.getSheet(sheetName);
	            if (sheet == null) {
	                sheet = workbook.createSheet(sheetName);
	            }

	            // Create or get header row
	            Row headerRow = sheet.getRow(0);
	            if (headerRow == null) {
	                headerRow = sheet.createRow(0);
	            }

	            // Set header titles for columns 7 and 8
	            Cell headerCellCol7 = headerRow.getCell(6); // Column G
	            if (headerCellCol7 == null || headerCellCol7.getCellType() == CellType.BLANK) {
	                headerCellCol7 = headerRow.createCell(6);
	                headerCellCol7.setCellValue("CTO Details");
	            }

	            Cell headerCellCol8 = headerRow.getCell(7); // Column H
	            if (headerCellCol8 == null || headerCellCol8.getCellType() == CellType.BLANK) {
	                headerCellCol8 = headerRow.createCell(7);
	                headerCellCol8.setCellValue("CTO Links");
	            }

	            // Create or get target row
	            Row row = sheet.getRow(iteration);
	            if (row == null) {
	                row = sheet.createRow(iteration);
	            }

	            // Write data to columns 7 and 8
	            Cell cellCol7 = row.createCell(6); // Column G
	            cellCol7.setCellValue(dataCol7);

	            Cell cellCol8 = row.createCell(7); // Column H
	            cellCol8.setCellValue(dataCol8);

	            fis.close(); // Close input stream before writing

	            // Save changes to file
	            try (FileOutputStream fos = new FileOutputStream(filePath)) {
	                workbook.write(fos);
	            }

	            workbook.close();
	            System.out.println("Data written to row " + iteration + " in columns 7 and 8.");

	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	public static void writeComInMasterSheet(String sheetName, int columnNumber, String data,String name, String designation, String companyName, String fileName,
			String sheetname, int iteration) throws IOException {
		File file = new File(System.getProperty("user.dir") + "//Master Company link Sheet.xlsx");
		Workbook workbook;
		Sheet sheet;

		// Load existing workbook or create new one
		if (file.exists()) {
			FileInputStream fileInputStream = new FileInputStream(file);
			workbook = new XSSFWorkbook(fileInputStream);
			fileInputStream.close();
		} else {
			workbook = new XSSFWorkbook();
		}

		// Get or create sheet
		sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			sheet = workbook.createSheet(sheetName);
		}

		// Check for duplicate in the target column
		boolean isDuplicate = false;
		int lastRowNum = sheet.getLastRowNum();
		for (int i = 0; i <= lastRowNum; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				Cell cell = row.getCell(columnNumber);
				if (cell != null && cell.getCellType() == CellType.STRING) {
					if (cell.getStringCellValue().equalsIgnoreCase(data)) {
						isDuplicate = true;
						System.out.println("Duplicate entry in the Mastersheet");
						logger.info("Duplicate entry in the Mastersheet");
						break;
					}
				}
			}
		}

		// Write the data if not duplicate
		if (!isDuplicate) {
			Row newRow = sheet.createRow(lastRowNum + 1);
			Cell newCell = newRow.createCell(columnNumber);
			newCell.setCellValue(data);
			System.out.println("Data Written in Mastersheet - " + data);
			logger.info("Data Written in Mastersheet - " + data);
	//		writeCompanyDatas(fileName, sheetname,iteration, name, designation, companyName,data);

		}

		// Write to file
		FileOutputStream fileOutputStream = new FileOutputStream(file);
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		workbook.close();
	}

	

	
	
	public static void writeData(String fileName, String sheetName, String data) {
	    String filePath = fileName;
	    String dataToWrite = data;

	    try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fis)) {

	        Sheet sheet = workbook.getSheet(sheetName);
	        if (sheet == null) {
	            sheet = workbook.createSheet(sheetName);
	        }

	        // Find the first empty row starting from index 0
	        int lastRowNum = sheet.getLastRowNum();
	        int writeRowIndex = 0;

	        for (int i = 0; i <= lastRowNum; i++) {
	            Row existingRow = sheet.getRow(i);
	            Cell existingCell = (existingRow != null) ? existingRow.getCell(0) : null;

	            if (existingRow == null || existingCell == null || existingCell.getCellType() == CellType.BLANK
	                    || (existingCell.getCellType() == CellType.STRING && existingCell.getStringCellValue().isEmpty())) {
	                writeRowIndex = i;
	                break;
	            } else {
	                writeRowIndex = lastRowNum + 1;
	            }
	        }

	        Row writeRow = sheet.getRow(writeRowIndex);
	        if (writeRow == null) {
	            writeRow = sheet.createRow(writeRowIndex);
	        }

	        Cell writeCell = writeRow.createCell(0); // First column
	        writeCell.setCellValue(dataToWrite);

	        // Save changes
	        try (FileOutputStream fos = new FileOutputStream(filePath)) {
	            workbook.write(fos);
	            System.out.println("Data written at row index: " + writeRowIndex);
	            logger.info("Data written at row index: " + writeRowIndex);
	        }

	    } catch (IOException e) {
	        e.printStackTrace();
	    }
	}
	private static boolean isCellEmpty(Cell cell) {
	    return cell == null ||
	           cell.getCellType() == CellType.BLANK ||
	           (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty());
	}
	public static void writeDatas(String fileName, String sheetName, String data1, String data2) {
	    String filePath = fileName;

	    try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fis)) {

	        Sheet sheet = workbook.getSheet(sheetName);
	        if (sheet == null) {
	            sheet = workbook.createSheet(sheetName);
	        }

	        // Find first empty row starting from row 0
	        int writeRowIndex = 0;
	        boolean foundEmpty = false;
	        int lastRowNum = sheet.getLastRowNum();

	        for (int i = 0; i <= lastRowNum; i++) {
	            Row row = sheet.getRow(i);
	            if (row == null || (isCellEmpty(row.getCell(0)) && isCellEmpty(row.getCell(1)))) {
	                writeRowIndex = i;
	                foundEmpty = true;
	                break;
	            }
	        }

	        if (!foundEmpty) {
	            writeRowIndex = lastRowNum + 1;
	        }

	        Row writeRow = sheet.getRow(writeRowIndex);
	        if (writeRow == null) {
	            writeRow = sheet.createRow(writeRowIndex);
	        }

	        // Write data to columns A and B (index 0 and 1)
	        writeRow.createCell(0).setCellValue(data1);
	        writeRow.createCell(1).setCellValue(data2);

	        // Optional: Auto-size columns
	        sheet.autoSizeColumn(0);
	        sheet.autoSizeColumn(1);

	        // Save changes
	        try (FileOutputStream fos = new FileOutputStream(filePath)) {
	            workbook.write(fos);
	            System.out.println("Data written at row index: " + writeRowIndex);
	        }

	    } catch (IOException e) {
	        e.printStackTrace();
	    }
	}

	// --------------------------------------------- Create New Excel File -----------------------------------------------------
    public static void createExcelFile(String folderName, String fileName) {
        String filePath = System.getProperty("user.dir") + "//" + folderName + "//" + fileName + ".xlsx";
        File file = new File(filePath);

        if (file.exists()) {
            System.out.println("Excel file already exists. Skipping creation.");
            logger.info(fileName + " - Excel file already exists. Skipping creation.");
            return;
        }

        // Create parent folder if not exists
        File folder = new File(System.getProperty("user.dir") + "//" + folderName);
        if (!folder.exists()) {
            folder.mkdirs();
        }

        Workbook workbook = new XSSFWorkbook();
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
            workbook.close();
            System.out.println("Excel file created successfully!");
            logger.info(fileName + " - Excel file created successfully!");
        } catch (IOException e) {
            e.printStackTrace();
            logger.warn("Error while creating Excel file: " + e.getMessage());
        }
    }
	
	/*public static void createExcelFile(String folderName,String fileName) {
		// Create a new workbook
		Workbook workbook = new XSSFWorkbook();
		// Sheet sheet;
		// Create a new sheet
		// sheet = workbook.createSheet(sheetName1);
		// sheet = workbook.createSheet(sheetName2);

		// Create a row
		// Row row = sheet.createRow(0);

		// Create cells
		// row.createCell(0).setCellValue("Links");
		// row.createCell(1).setCellValue("Password");

		// Add data to next row
		// Row dataRow = sheet.createRow(1);
		// dataRow.createCell(0).setCellValue(links);
		// dataRow.createCell(1).setCellValue("admin123");

		// Write to file
		try (FileOutputStream fileOut = new FileOutputStream(System.getProperty("user.dir")+"//"+folderName+"//"+fileName + ".xlsx")) {
			workbook.write(fileOut);
			workbook.close();
			System.out.println("Excel file created successfully!");
			logger.info(fileName+" - Excel file created successfully!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}*/

	public static void writeDatainExcel(String sheetName, int rownum, int cellnum, String data) throws IOException {
		File file = new File(System.getProperty("user.dir") + "//" + stamp + "- Unique Links.xlsx");
		// Open the file input stream
		FileInputStream fileInputStream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(fileInputStream);

		// Get the sheet
		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			sheet = workbook.createSheet(sheetName);
		}

		// Get or create the row
		Row row = sheet.getRow(rownum);
		if (row == null) {
			row = sheet.createRow(rownum);
		}

		// Create the cell and set value
		Cell cell = row.createCell(cellnum);
		cell.setCellValue(data);

		// Close the input stream before writing
		fileInputStream.close();

		// Write to the file
		FileOutputStream fileOutputStream = new FileOutputStream(file);
		workbook.write(fileOutputStream);

		// Close resources
		fileOutputStream.close();
		workbook.close();
	}
	
	
	

    public int getFirstEmptyRowIndexIn8Columns(String filePath, String sheetName) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            int lastRowNum = sheet.getLastRowNum();

            for (int i = 0; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                if (isFirst8ColumnsEmpty(row)) {
                    return i;
                }
            }

            // No empty row found in existing rows, return next available row
            return lastRowNum + 1;
        }
    }

    private static boolean isFirst8ColumnsEmpty(Row row) {
        if (row == null) return true;

        for (int col = 0; col < 8; col++) {
            Cell cell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if (cell.getCellType() != CellType.BLANK &&
                cell.getCellType() != CellType._NONE &&
                !cell.toString().trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }


    public void countFilledCellsInExcel(String filepath) {
        String excelFilePath = filepath; // Update with your actual file path
        int filledCellCount = 0;

        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // You can loop through multiple sheets if needed

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell != null && cell.getCellType() != CellType.BLANK &&
                        !cell.toString().trim().isEmpty()) {
                        filledCellCount++;
                    }
                }
            }

            System.out.println("Total filtered Link count: " + filledCellCount);
            logger.info("Total filtered Link count: " + filledCellCount);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void autoFitColumn(String filePath, String sheetName) throws IOException {
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(sheetName);

        if (sheet != null) {
            // Loop through all columns in the first row
            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                int totalColumns = headerRow.getLastCellNum(); // total columns
                for (int i = 0; i < totalColumns; i++) {
                    sheet.autoSizeColumn(i); // Auto-fit for each column
                }
            }
        }

        fis.close();
        FileOutputStream fos = new FileOutputStream(filePath);
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("Column widths auto-fitted successfully!");
    }
//----------------------------------------------------------------Log File Appender -------------------------------------------------------
	   public static  Logger logger= Logger.getLogger(Google_Search_002.class);
	   static {
			    String stamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm").format(new Date());
		        String logFileName = "logs/" + stamp + "-Company filter Log.log";

		        try {
		            RollingFileAppender rollingFileAppender = new RollingFileAppender();
		            rollingFileAppender.setName("FileLogger");
		            rollingFileAppender.setFile(logFileName);
		            rollingFileAppender.setMaxFileSize("5MB");
		            rollingFileAppender.setMaxBackupIndex(10);
		            rollingFileAppender.setLayout(new PatternLayout("%d{ISO8601} %-5p %c{1} - %m%n"));
		            rollingFileAppender.activateOptions();
		            logger.addAppender(rollingFileAppender);
		        } catch (Exception e) {
		            e.printStackTrace();
		        }
	   }  
}
