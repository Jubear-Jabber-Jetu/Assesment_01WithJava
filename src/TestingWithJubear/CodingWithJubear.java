package TestingWithJubear;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.time.Duration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class CodingWithJubear {
    public static void main(String[] args) {
        // Set the path to the FirefoxDriver executable
        System.setProperty("webdriver.gecko.driver", "C:\\Users\\Walton\\OneDrive\\Desktop\\geckodriver-v0.33.0-win64\\geckodriver.exe");

        // Create FirefoxOptions
        FirefoxOptions options = new FirefoxOptions();

        // Create a new instance of the FirefoxDriver with the specified options
        WebDriver driver = new FirefoxDriver(options);

        // Set implicit wait time
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

        // Open Google in Firefox
        driver.get("https://www.google.com");

        // Specify the path to your Excel file
        String excelFilePath = "C:\\Users\\Walton\\OneDrive\\Desktop\\Excel.xlsx";

        try {
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);

            // Get the current day name
            String currentDayName = new SimpleDateFormat("EEEE").format(new Date());

            // Use the current day name as the worksheet name
            Sheet worksheet = workbook.getSheet(currentDayName);
            if (worksheet == null) {
                System.out.println("Worksheet '" + currentDayName + "' does not exist. Exiting.");
                return;
            }

            // Iterate through rows and get non-empty values from the 3rd column (column 'C')
            for (Row row : worksheet) {
                Cell cell = row.getCell(2); // Column 'C'
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String searchTerm = cell.getStringCellValue();
                    if (!searchTerm.isEmpty()) {
                        // Enter the search term into the Google search bar
                        WebElement searchBox = driver.findElement(By.name("q"));
                        searchBox.clear();
                        searchBox.sendKeys(searchTerm);

                        // Wait for suggestions to appear
                        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//ul[@role='listbox']/li[@role='presentation']")));

                        // Get the suggestions from the search bar
                        List<WebElement> suggestions = driver.findElements(By.xpath("//ul[@role='listbox']/li[@role='presentation']"));

                        // Debugging: Print the keyword and its suggestions
                        System.out.println("Keyword: " + searchTerm);
                        System.out.println("Suggestions: ");
                        for (WebElement suggestion : suggestions) {
                            System.out.println(suggestion.getText());
                        }

                        // Initialize variables to store the longest and shortest suggestions for the current keyword
                        String maxSuggestion = "";
                        String minSuggestion = suggestions.get(0).getText();

                        // Find the suggestion with the maximum and minimum text length
                        for (WebElement suggestion : suggestions) {
                            String text = suggestion.getText();
                            if (text.length() > maxSuggestion.length()) {
                                maxSuggestion = text;
                            }
                            if (text.length() < minSuggestion.length()) {
                                minSuggestion = text;
                            }
                        }

                        // Truncate suggestions to fit in one cell if they are too long
                        maxSuggestion = truncateText(maxSuggestion, 32767); // Excel cell limit
                        minSuggestion = truncateText(minSuggestion, 32767); // Excel cell limit

                        // Update the Excel sheet with the Longest and Shortest Options for the current keyword
                        Cell longestCell = row.createCell(3); // Column 'D' for Longest Option
                        longestCell.setCellValue(maxSuggestion);

                        Cell shortestCell = row.createCell(4); // Column 'E' for Shortest Option
                        shortestCell.setCellValue(minSuggestion);
                    }
                }
            }

            // Save the updated Excel workbook
            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            System.out.println("Process completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Close the WebDriver when done
            driver.quit();
        }
    }

    private static String truncateText(String text, int maxLength) {
        if (text.length() > maxLength) {
            return text.substring(0, maxLength);
        }
        return text;
    }
}