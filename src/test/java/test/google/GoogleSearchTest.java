package test.google;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.*;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.List;

public class GoogleSearchTest extends BaseTest {
    private static final By SEARCH_INPUT = By.id("APjFqb");
    private static final By GOOGLE_SEARCH_LIST = By.xpath("//li[contains(@class, 'sbct')]//div[@class='wM6W7d']/span");

    @Test
    public void unitTwoFinalTask() throws IOException, InterruptedException {

        List<WebElement> elements = driver.findElements(SEARCH_INPUT);
        Assert.assertFalse(elements.isEmpty(), "Home Page is not Displayed");
        WebElement inputField = driver.findElement(SEARCH_INPUT);

        // Finding Today Day From System
        DayOfWeek today = LocalDate.now().getDayOfWeek();
        String todayDayName= today.toString().substring(0, 1) + today.toString().substring(1).toLowerCase();

        //Load Excel file
        InputStream excel = getClass().getClassLoader().getResourceAsStream("Excel.xlsx");
        if (excel == null) {
            throw new FileNotFoundException("Excel.xlsx file is missing in the resources folder");
        }
        Workbook workbook = new XSSFWorkbook(excel);
        Sheet sheet = workbook.getSheet(todayDayName);


        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Cell keywordCell = row.getCell(2); // Get Input Key Value in 3rd column
            if (keywordCell != null) {
                String keyword = keywordCell.getStringCellValue();

                // Perform Google search
                inputField.clear();
                inputField.sendKeys(keyword);

                // Wait for search suggestions to load
                Thread.sleep(2000); // Adjust as necessary

                // Extract search suggestions
                List<WebElement> suggestions = driver.findElements(GOOGLE_SEARCH_LIST);
                String longest = "", shortest = "";
                for (WebElement suggestion : suggestions) {
                    String text = suggestion.getText();
                    String trimText = suggestion.getText().trim();
                    if (trimText.isEmpty()) continue;
                    if (trimText.length() > longest.length())
                        longest = text;
                    if (shortest.isEmpty() || trimText.length() < shortest.length())
                        shortest = text;
                }

                // Write results back to Excel
                Cell longestCell = row.createCell(3); // 4th column
                longestCell.setCellValue(longest);

                Cell shortestCell = row.createCell(4); // 5th column
                shortestCell.setCellValue(shortest);
            }
        }
        // Save updated Excel file
        String downloadsDirectory = System.getProperty("user.home") + File.separator + "Downloads";
        File outputFile = new File(downloadsDirectory + File.separator + "Updated_Excel.xlsx");
        File directory = new File(downloadsDirectory);
        if (!directory.exists()) {
            directory.mkdirs();
        }

        FileOutputStream fos = new FileOutputStream(outputFile);
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("Excel file updated and saved successfully at: " + outputFile.getAbsolutePath());
    }

}
