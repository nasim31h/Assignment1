import java.io.IOException;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class GoogleSearch {

    private static WebDriver driver;

    public static void main(String[] args) throws IOException {
        // Set up the Chrome driver
        System.setProperty("webdriver.chrome.driver", "/path/to/chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        options.setHeadless(false);
        driver = new ChromeDriver(options);

        // Get the Excel file
        String file_path = "Excel.xlsx";
        readLoadData(file_path);

        // Quit the driver
        driver.quit();
    }

    private static void readLoadData(String file_path) throws IOException {
        // Load the Excel file
        ExcelUtils.loadExcelFile(file_path);

        // Iterate through all the sheets
        for (String sheetName : ExcelUtils.getSheetNames()) {
            // Get the keywords to search
            List<String> keywords = ExcelUtils.getKeywords(sheetName);

            // Start iterating from the 3rd row
            int startRow = 3;

            // Search for each keyword and print the results
            for (String keyword : keywords) {
                searchGoogle(keyword, startRow, sheetName);
                startRow++;
            }
        }
    }

    private static void searchGoogle(String keyword, int startRow, String sheetName) {
        // Navigate to the Google search page
        driver.get("https://www.google.com/");

        // Find the search box element and enter keywords
        WebElement searchBox = driver.findElement(By.id("q"));
        searchBox.sendKeys(keyword);

        // Press the Enter key to submit the search
        searchBox.sendKeys(Keys.ENTER);

        // Wait for the search results to load
        WebDriverWait wait = new WebDriverWait(driver, 10);
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("g"));

        // Get the first and last items in the search results
        List<WebElement> searchResults = driver.findElements(By.className("g"));
        String firstItem = searchResults.get(0).getText();
        String lastItem = searchResults.get(searchResults.size() - 1).getText();

        // Save the first and last items to the Excel file
        ExcelUtils.saveFirstAndLastItems(startRow, firstItem, lastItem, sheetName);

        // Get the longest and shortest option item
        String longestOptionItem = searchResults.stream().max(String::length).get();
        String shortestOptionItem = searchResults.stream().min(String::length).get();

        // Save the longest and shortest option item to the Excel file
        ExcelUtils.saveLongestShortestOptionItem(startRow + 2, longestOptionItem, shortestOptionItem, sheetName);
    }
}

class ExcelUtils {

    public static void loadExcelFile(String file_path) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(file_path);
    }

    public static List<String> getKeywords(String sheetName) {
        List<String> keywords = new ArrayList<>();
        XSSFWorkbook workbook = new XSSFWorkbook();
        int sheetIndex = workbook.getSheetIndex(sheetName);
        if (sheetIndex == -1) {
            return keywords;
        }
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(sheetIndex);
        int rowCount = sheet.getLastRowNum() + 1;
        for (int i = 2; i < rowCount; i++) {
            String keyword = sheet.getRow(i).getCell(0).getStringCellValue();
            
         keywords.add(keyword);
          }
        return keywords;
    }

    public static void saveFirstAndLastItems(int startRow, String firstItem, String lastItem, String sheetName) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        int sheetIndex = workbook.getSheetIndex(sheetName);
        if (sheetIndex == -1) {
            return;
        }
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(sheetIndex);
        sheet.getRow(startRow).getCell(0).setCellValue(firstItem);
        sheet.getRow(startRow + 1).getCell(0).setCellValue(lastItem);
        workbook.save(file_path);
    }

    // Added this method

    public static void saveLongestShortestOptionItem(int startRow, String longestOptionItem, String shortestOptionItem, String sheetName) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        int sheetIndex = workbook.getSheetIndex(sheetName);
        if (sheetIndex == -1) {
            return;
        }
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(sheetIndex);
        sheet.getRow(startRow + 2).getCell(0).setCellValue(longestOptionItem);
        sheet.getRow(startRow + 3).getCell(0).setCellValue(shortestOptionItem);
        workbook.save(file_path);
    }
}
