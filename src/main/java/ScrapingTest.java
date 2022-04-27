import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.*;

public class ScrapingTest extends Job {
    //the following list stores Job class objects
    private static List<Job> jobs = new ArrayList<>();

    public ScrapingTest(String name, String company, String location) {
        super(name, company, location);
    }
    //this method reads job offers from a page
    static void findJobs(WebDriver driver) throws IOException {

        int numberOfJobsOnPage = driver.findElements(By.className("jobTitle")).size();
        List<WebElement> jobNames = driver.findElements(By.className("jobTitle"));
        List<WebElement> jobCompany = driver.findElements(By.className("companyName"));
        List<WebElement> jobLocation = driver.findElements(By.className("companyLocation"));

        List<String> jobNamesToString = new ArrayList<>();
        List<String> jobCompanyToString = new ArrayList<>();
        List<String> jobLocationToString = new ArrayList<>();

        for (WebElement selectedJob : jobNames) {
            jobNamesToString.add(selectedJob.getText());
        }
        for (WebElement selectedCompany : jobCompany) {
            jobCompanyToString.add(selectedCompany.getText());
        }
        for (WebElement selectedLocation : jobLocation) {
            jobLocationToString.add(selectedLocation.getText());
        }

        new Job("", "", "");

        for (int i = 0; i < numberOfJobsOnPage; i++) {
            jobs.add(new Job(jobNamesToString.get(i), jobCompanyToString.get(i), jobLocationToString.get(i)));
            //System.out.println(jobs.get(i));
        }
    }
    //this method writes job offers into Excel file
    static void writeJobs() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet(" Job details ");
        XSSFRow row;

        Map<Integer, Object[]> data
                = new TreeMap<>();
        
        for (int i = 0; i < jobs.size(); i++) {
            data.put(i, new Object[]{jobs.get(i)});
        }

        Set<Integer> keyid = data.keySet();
        int rowid = 0;

        for (Integer key : keyid) {
            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = data.get(key);
            int cellid = 0;

            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue(obj.toString());
            }
        }
        FileOutputStream out = new FileOutputStream(
                new File("C:/Users/adam/Desktop/JobOffers.xlsx"));

        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) throws InterruptedException {

        System.setProperty("webdriver.gecko.driver", "C:\\Users\\adam\\Desktop\\untitled\\geckodriver-v0.30.0-win64\\geckodriver.exe");
        WebDriver driver = new FirefoxDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        driver.navigate().to("https://pl.indeed.com/?r=us");
        driver.manage().window().maximize();
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("onetrust-reject-all-handler"))).click();

        Scanner scan = new Scanner(System.in);
        System.out.println("Select job: ");
        String job = scan.next();
        System.out.println("Select location: ");
        String place = scan.next();

        WebElement jobTextBox = driver.findElement(By.id("text-input-what"));
        WebElement placeTextBox = driver.findElement(By.id("text-input-where"));

        //searching for job offers with details given by user
        jobTextBox.click();
        jobTextBox.sendKeys(job);
        placeTextBox.click();
        placeTextBox.sendKeys(place);

        WebElement searchButton = driver.findElement(By.className("yosegi-InlineWhatWhere-primaryButton"));
        searchButton.click();

        //checking if given job details are correct
        try{
            if(driver.findElement(By.cssSelector("#resultsCol > div.no_results > div > h1")).isDisplayed() || driver.findElement(By.cssSelector(".bad_query > h1:nth-child(1)")).isDisplayed()){
            System.out.println("Job or location was misspelled");
            driver.quit();
            System.exit(0);
            }
        } catch (NoSuchElementException exception){
            }
        try {
            List<WebElement> pagination = driver.findElements(By.xpath("//ul[@class='pagination-list']/li"));
            int pageSize = pagination.size();

            findJobs(driver);
            writeJobs();
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("button.icl-CloseButton:nth-child(3)"))).click();

            Thread.sleep(1000);
            driver.findElement(By.className("pn")).click();
            Thread.sleep(1000);
            driver.findElement(By.className("popover-x-button-close")).click();
            findJobs(driver);

            for (int i = 4; i <= pageSize - 1; i++) {
                WebElement pageNum = driver.findElement(By.xpath("(//ul[@class='pagination-list']/li)[" + i + "]"));
                pageNum.click();
                findJobs(driver);

                //handling 5th page which has different structure
                if (i == pageSize - 1) {
                    driver.findElement(By.cssSelector("#resultsCol > nav > div > ul > li:nth-child(5) > a > span")).click();
                    findJobs(driver);
                }
            }
           writeJobs();
           System.out.println("Job offers have been downloaded");
           driver.quit();
           System.exit(0);

        } catch (NoSuchElementException | IOException exception) {
            driver.quit();
            System.out.println("Job offers have been downloaded");
            System.exit(0);
        }
    }
}


