package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

public class FirefoxSearchSuggestions {

    public static void main(String[] args) throws InterruptedException, IOException {
        // Set up WebDriverManager to automatically manage FirefoxDriver
        WebDriverManager.firefoxdriver().setup();

        // Set Firefox options (optional)
        FirefoxOptions options = new FirefoxOptions();
        options.setHeadless(false); // Set to true if you want headless mode (no GUI)

        // Initialize the WebDriver
        WebDriver driver = new FirefoxDriver(options);


        try {
            // Open the URL (e.g., Google)
            File file = new File("search_suggestions.xlsx");  // Replace with your Excel file
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);
            int numberOfSheets = workbook.getNumberOfSheets();

            LocalDate today = LocalDate.now();
        
            // Get the day of the week (DayOfWeek is an enum)
            DayOfWeek dayOfWeek = today.getDayOfWeek();
            
            // Get the name of the day (e.g., MONDAY, TUESDAY, etc.)
            String dayName = dayOfWeek.name();

            for(int i=0;i<numberOfSheets;i++){
                Sheet sheet = workbook.getSheetAt(i);
                if(sheet.getSheetName().equals(dayName)){

                    for (int rowIndex = 1; rowIndex < 8; rowIndex++) {
                        Row row = sheet.getRow(rowIndex);
                        
                        String keyword = row.getCell(1).getStringCellValue();
            
                        // Perform Google search
                        driver.get("https://www.google.com");

                        // Wait until the search box is present
                        WebDriverWait wait = new WebDriverWait(driver, 10);
                        WebElement searchBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("q")));
                        // WebElement searchBox = driver.findElement(By.name("q"));
                    // Type a search query (e.g., "Selenium")
                        searchBox.sendKeys(keyword);

                        // Wait for search suggestions to appear
                        List<WebElement> suggestions = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("ul[role='listbox'] li")));
                        //List<WebElement> suggestions = driver.findElements(By.cssSelector("ul[role='listbox'] li"));
                        if (!suggestions.isEmpty()) {
                            String longestSuggestion = "";
                            String shortestSuggestion = suggestions.get(0).getText();
            
                            for (WebElement suggestion : suggestions) {

                                String suggestionText = suggestion.getText();
                                System.out.println("Get tex AAA: "+suggestionText);

                                if (suggestionText.length() > longestSuggestion.length()) {
                                    longestSuggestion = suggestionText;
                                }
                                if (suggestionText.length() < shortestSuggestion.length()) {
                                    shortestSuggestion = suggestionText;
                                }
                            }
            
                            // Write the longest and shortest suggestions back to the same Excel file


                        // Write values to new row         // Column A - Keyword
                            row.createCell(3).setCellValue(longestSuggestion); // Column B - Longest Suggestion
                            row.createCell(2).setCellValue(shortestSuggestion); 
                        }
                    }
                }
            }
            
            FileOutputStream fos = new FileOutputStream(file);
            workbook.write(fos);
            fos.close();
        
            // Close the Excel workbook
            workbook.close();
            System.out.println("Search suggestions written to 'search_suggestions.xlsx'");

        } finally {
            // Close the WebDriver
            driver.quit();
        }
    }
}
