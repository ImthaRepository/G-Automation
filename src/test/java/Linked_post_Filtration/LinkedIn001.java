package Linked_post_Filtration;

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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;


public class LinkedIn001 {

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
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("username")));
		
		WebElement usernameField = driver.findElement(By.id("username"));
		usernameField.clear();
		usernameField.sendKeys(getPropertyFileValue("username"));
		
		
		WebElement passwordField = driver.findElement(By.id("password"));
		passwordField.clear();
		passwordField.sendKeys(getPropertyFileValue("password"));
		
		driver.findElement(By.xpath("//button[@aria-label='Sign in']")).click();

	}

	@AfterTest
	public void signOut() {
		driver.findElement(By.xpath("//img[contains(@class,'global-nav__me-photo')]")).click();
		System.out.println("Profile Icon clicked");
		WebElement signOutBtn = driver.findElement(By.xpath("//a[text()= 'Sign Out']"));
		JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].click();", signOutBtn);
		System.out.println("Sign out button clicked");
		try {
			driver.findElement(By.xpath("//button//span[text()= 'Sign out']")).click();
			System.out.println("Sign out in the pop up is clicked");
		} catch (Exception e) {
			System.out.println("Sign out pop up not available");
		}
		
		
	}
	
	@AfterSuite
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

		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Search']")));

		WebElement searchfield = driver.findElement(By.xpath("//input[@placeholder='Search']"));
		searchfield.sendKeys(getPropertyFileValue("keyword"), Keys.ENTER);

		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//button[text()='Posts'])[1]")));
		driver.findElement(By.xpath("(//button[text()='Posts'])[1]")).click();
//		
//		driver.findElement(By.xpath("//button[text()='Date posted']")).click();
//		driver.findElement(By.xpath("//span[text()='Past week']")).click(); Past 24 hours
//		
//		driver.findElement(By.xpath("(//span[text()='Show results'])[1]")).click();
//		
		JavascriptExecutor js = (JavascriptExecutor) driver;
		long lastHeight = (long) js.executeScript("return document.body.scrollHeight");
		int scrollPauseTime = 3000;
		String loopCount = getPropertyFileValue("ScrollCount");
		System.out.println("The Number of Scroll count Entered is - " + loopCount);
		logger.info("The Number of Scroll count Entered is - " + loopCount);
		int ScrollCount = Integer.parseInt(loopCount);
		for (int i = 0; i <= ScrollCount; i++) {
			System.out.println("Scroll: " + i);
			logger.info("Scroll: " + i);
			js.executeScript("window.scrollTo(0, document.body.scrollHeight);");

			// Wait for new content to load
			Thread.sleep(scrollPauseTime); // adjust based on site loading speed
			js.executeScript("window.scrollTo( document.body.scrollHeight, (document.body.scrollHeight)/2);");
			Thread.sleep(3000);
			long newHeight = (long) js.executeScript("return document.body.scrollHeight");
			lastHeight = newHeight;

		}

		createExcelFile("excel files Raw", stamp);
		int linkCount = driver.findElements(By.xpath("//span[@class='update-components-actor__title']")).size();
		System.out.println("Total link is " + linkCount);
		logger.info("Total link is " + linkCount);
		for (int i = 1; i <= linkCount; i++) {
			try {
				driver.findElement(By.xpath("(//span[@class='update-components-actor__title'])["+i+"]/ancestor::div[@class='fie-impression-container']//span[text()='View job preferences']"));
				System.out.println(i+"- The post is open to work");
				logger.info(i+"- The post is open to work");
			} catch (Exception e) {
				try {
					WebElement linkTag = driver.findElement(
							By.xpath("(//span[@class='update-components-actor__title'])[" + i + "]/ancestor::a"));
					String href = linkTag.getAttribute("href");
					try {
						insertValueCell("All_Links", i - 1, 0, href);
						System.out.println(i + " - Data written successfully.");
						logger.info(i + " - Data written successfully.");
					} catch (IOException b) {
						b.printStackTrace();
						System.out.println(i + " - Failed to Write Data");
						logger.error(i + " - Failed to Write Data");
					}
				} catch (Exception c) {
					c.printStackTrace();
					System.out.println(i + " - Link is disabled");
					logger.warn(i + " - Link is disabled");
				}
			}

		}
		deleteExcelIfNoSheets();
		createExcelFile("excel files filtered",stamp+"-Filtered");
	}

	// -------------------------------------------------Fetched link Validation and Filter----------------------------------------------
	static int iteration = 0;
	static int iterate = 0;

	@Test(priority = 2, dataProvider = "excelData", dependsOnMethods = "Fetchdata")
	public void writeURLs(String URL1) throws IOException {

		String filepath = System.getProperty("user.dir") + "\\excel files filtered\\" + stamp + "-Filtered.xlsx";
		try {
			writeInMasterSheet("Links", 0, URL1, filepath, "Sheet1");
			System.out.println(iterate + " - link written in Master sheet");
			logger.info(iterate + " - link written in master sheet");
		} catch (Exception e) {
			System.out.println(iterate + " - link failed to written in Master sheet");
			logger.info(iterate + " - link failed to written in master sheet");
		}
		iterate++;
	}
	
	
	//----------------------------------------------------Delete Excel file if it has no Sheet -----------------------------------------------	
		@Test(priority = 3, dependsOnMethods = "writeURLs")
		public void deleteExcelIfNoSheet() throws FileNotFoundException, IOException {
		
			String filePath = System.getProperty("user.dir") + "\\excel files filtered\\" + stamp + "-Filtered.xlsx"; // Replace with actual file path
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
				}

			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		
		public void deleteExcelIfNoSheets() throws FileNotFoundException, IOException {
			String filePath = System.getProperty("user.dir") + "\\excel files Raw\\" + stamp +".xlsx"; // Replace with actual file path
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
	public Object[][] getDataFromExcel() {
		String filePath = System.getProperty("user.dir") + "\\excel files Raw\\" + stamp + ".xlsx";
		String sheetName = "All_Links";
		return readExcelData(filePath, sheetName);
	}

	private Object[][] readExcelData(String filePath, String sheetName) {
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

//-----------------------------------------------------------Number Conversion Code------------------------------------------------------------//

	public static long convertToNumber(String value) {
		value = value.trim().toUpperCase();

		if (value.endsWith("K")) {
			return (long) (Double.parseDouble(value.replace("K", "")) * 1_000);
		} else if (value.endsWith("K+")) {
			return (long) (Double.parseDouble(value.replace("K+", "")) * 1_000);
		} else if (value.endsWith("M")) {
			return (long) (Double.parseDouble(value.replace("M", "")) * 1_000_000);
		} else if (value.endsWith("M+")) {
			return (long) (Double.parseDouble(value.replace("M+", "")) * 1_000_000);
		} else if (value.endsWith("+")) {
			return (long) (Double.parseDouble(value.replace("+", "")) * 1);
		} else if (value.endsWith("B")) {
			return (long) (Double.parseDouble(value.replace("B", "")) * 1_000_000_000);
		} else {
			return Long.parseLong(value); // if it's already a number
		}
	}

	public static int convertReviewToNumber(String value) {
		value = value.trim().toUpperCase();

		if (value.endsWith("K reviews") || value.endsWith("K+ reviews")) {
			return (int) (Double.parseDouble(value.replaceAll("[^\\d.]", "")) * 1_000);
		} else if (value.endsWith("M reviews") || value.endsWith("M+ reviews")) {
			return (int) (Double.parseDouble(value.replaceAll("[^\\d.]", "")) * 1_000_000);
		} else if (value.endsWith("B reviews") || value.endsWith("B+ reviews")) {
			return (int) (Double.parseDouble(value.replaceAll("[^\\d.]", "")) * 1_000_000_000);
		} else if (value.endsWith("+ reviews") || value.endsWith(" reviews")) {
			return (int) Double.parseDouble(value.replaceAll("[^\\d.]", ""));
		} else {
			return Integer.parseInt(value.replaceAll("[^\\d]", "")); // fallback
		}
	}

//------------------------------------------------------------ Insert Value in the Excel Sheet --------------------------------------------//

	public static void insertValueCell(String sheetName, int rownum, int cellnum, String data) throws IOException {
		// File file = new File(System.getProperty("user.dir") + "//playLinks.xlsx");
		File file = new File(System.getProperty("user.dir") + "//excel files Raw//" + stamp + ".xlsx");
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

	public static void insertValueCellWithoutDuplicate(String sheetName, int rownum, int cellnum, String data)
			throws IOException {
		File file = new File(System.getProperty("user.dir") + "//excel files Raw//" + stamp + ".xlsx");

		Workbook workbook;
		Sheet sheet;

		// If file exists, read it; else create new workbook
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

		// Check for duplicates in the entire column (cellnum)
		boolean isDuplicate = false;
		for (Row existingRow : sheet) {
			Cell existingCell = existingRow.getCell(cellnum);
			if (existingCell != null && existingCell.getCellType() == CellType.STRING) {
				if (existingCell.getStringCellValue().equalsIgnoreCase(data)) {
					isDuplicate = true;
					break;
				}
			}
		}

		if (!isDuplicate) {
			// Get or create the row
			Row row = sheet.getRow(rownum);
			if (row == null) {
				row = sheet.createRow(rownum);
			}

			// Create the cell and set value
			Cell cell = row.createCell(cellnum);
			cell.setCellValue(data);

			// Write to file
			FileOutputStream fileOutputStream = new FileOutputStream(file);
			workbook.write(fileOutputStream);
			fileOutputStream.close();
			System.out.println("Data written successfully: " + data);
			logger.info("Data written successfully: " + data);
		} else {
			System.out.println("Duplicate entry skipped: " + data);
			logger.info("Duplicate entry skipped: " + data);
		}

		workbook.close();
	}

//	public static void writeInMasterSheet(String sheetName, int columnNumber, String data) throws IOException {
//		File file = new File(System.getProperty("user.dir") + "//MasterSheet.xlsx");
//		Workbook workbook;
//		Sheet sheet;
//
//		// Load existing workbook or create new one
//		if (file.exists()) {
//			FileInputStream fileInputStream = new FileInputStream(file);
//			workbook = new XSSFWorkbook(fileInputStream);
//			fileInputStream.close();
//		} else {
//			workbook = new XSSFWorkbook();
//		}
//
//		// Get or create sheet
//		sheet = workbook.getSheet(sheetName);
//		if (sheet == null) {
//			sheet = workbook.createSheet(sheetName);
//		}
//
//		// Check for duplicate in the target column
//		boolean isDuplicate = false;
//		int lastRowNum = sheet.getLastRowNum();
//		for (int i = 0; i <= lastRowNum; i++) {
//			Row row = sheet.getRow(i);
//			if (row != null) {
//				Cell cell = row.getCell(columnNumber);
//				if (cell != null && cell.getCellType() == CellType.STRING) {
//					if (cell.getStringCellValue().equalsIgnoreCase(data)) {
//						isDuplicate = true;
//						break;
//					}
//				}
//			}
//		}
//	}

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

	public static void writeData(String fileName, String sheetName, String data) {
		String filePath = fileName;
		String dataToWrite = data;

		try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fis)) {

			Sheet sheet = workbook.getSheet(sheetName);
			if (sheet == null) {
				sheet = workbook.createSheet(sheetName);
			}

			int lastRowNum = sheet.getLastRowNum();
			int writeRowIndex = 0;

			// Find the first empty row
			for (int i = 0; i <= lastRowNum; i++) {
				Row existingRow = sheet.getRow(i);
				Cell existingCell = (existingRow != null) ? existingRow.getCell(0) : null;

				if (existingRow == null || existingCell == null || existingCell.getCellType() == CellType.BLANK
						|| (existingCell.getCellType() == CellType.STRING
								&& existingCell.getStringCellValue().isEmpty())) {
					writeRowIndex = i;
					break;
				} else {
					writeRowIndex = lastRowNum + 1; // All filled, write on a new row
				}
			}

			Row writeRow = sheet.getRow(writeRowIndex);
			if (writeRow == null) {
				writeRow = sheet.createRow(writeRowIndex);
			}

			Cell writeCell = writeRow.createCell(0); // Always write in first column
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
	
	
	public void compareData(String sheetName, int columnNumber, String data) throws IOException {

		String masterFilePath = System.getProperty("user.dir") + "//Master Sheet.xlsx";
		String newFilePath = System.getProperty("user.dir") + "//" + stamp + "- Unique Links.xlsx";

		boolean isDuplicate = false;

		// Load Master Sheet
		FileInputStream masterInputStream = new FileInputStream(masterFilePath);
		Workbook masterWorkbook = new XSSFWorkbook(masterInputStream);
		Sheet masterSheet = masterWorkbook.getSheet(sheetName);

		if (masterSheet != null) {
			int lastRowNum = masterSheet.getLastRowNum();
			for (int i = 0; i <= lastRowNum; i++) {
				Row row = masterSheet.getRow(i);
				if (row != null) {
					Cell cell = row.getCell(columnNumber);
					if (cell != null && cell.getCellType() == CellType.STRING) {
						if (cell.getStringCellValue().equalsIgnoreCase(data)) {
							isDuplicate = true;
							break;
						}
					}
				}
			}
		}

		masterWorkbook.close();
		masterInputStream.close();

		// If not duplicate, write to new file
		if (!isDuplicate) {
			File newFile = new File(newFilePath);
			Workbook newWorkbook;
			Sheet newSheet;

			if (newFile.exists()) {
				FileInputStream newFileInput = new FileInputStream(newFile);
				newWorkbook = new XSSFWorkbook(newFileInput);
				newFileInput.close();
			} else {
				newWorkbook = new XSSFWorkbook();
			}

			newSheet = newWorkbook.getSheet(sheetName);
			if (newSheet == null) {
				newSheet = newWorkbook.createSheet(sheetName);
			}

			int newLastRow = newSheet.getLastRowNum();
			Row newRow = newSheet.createRow(newLastRow + 1);
			Cell newCell = newRow.createCell(columnNumber);
			newCell.setCellValue(data);

			FileOutputStream outputStream = new FileOutputStream(newFile);
			newWorkbook.write(outputStream);
			outputStream.close();
			newWorkbook.close();
		} else {
			System.out.println("Data already exists in Master Sheet.");
			logger.info("Data already exists in Master Sheet.");
		}

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

//----------------------------------------------------------------Log File Appender -------------------------------------------------------
	   public static  Logger logger= Logger.getLogger(LinkedIn001.class);
	   static {
			    String stamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm").format(new Date());
		        String logFileName = "logs/" + stamp + "-Log.log";

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
