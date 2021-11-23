package daurora.nebula.app;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;


public class AppApplicationTests {

	WebDriver driver;

	public static final String USERNAME = System.getenv("josephdaurora_QWyv4q");
	public static final String AUTOMATE_KEY = System.getenv("cErBQMD1RFawwQsRq5nA");
	public static final String URL = "https://" + USERNAME + ":" + AUTOMATE_KEY + "@hub-cloud.browserstack.com/wd/hub";

	@Before
	public void setup() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();

	}

	@Test
	public void smarterTest() throws InterruptedException {
		driver.get("https://nebula.josephdaurora.com");
		WebDriverWait wait = new WebDriverWait(driver, 10);
		WebElement getStartedButton;
		getStartedButton = wait.until(ExpectedConditions.presenceOfElementLocated(new By.ByXPath("//*[@id=\"welcome_section\"]/div/a")));
		assertTrue(getStartedButton.isDisplayed());
		getStartedButton.click();
		WebElement notFormattedButton;
		notFormattedButton = wait.until(ExpectedConditions.presenceOfElementLocated(new By.ByXPath("//*[@id=\"getting-started-buttons\"]/table/tbody/tr/td[4]/div/a")));
		notFormattedButton.click();
		WebElement assignmentName, numStudents, numQuestions;
		assignmentName = wait.until(ExpectedConditions.presenceOfElementLocated(new By.ByXPath("/html/body/form/div[1]/input[1]")));
		numStudents = wait.until(ExpectedConditions.presenceOfElementLocated(new By.ByXPath("/html/body/form/div[1]/input[2]")));
		numQuestions = wait.until(ExpectedConditions.presenceOfElementLocated(new By.ByXPath("/html/body/form/div[1]/input[3]")));
		assignmentName.sendKeys("Assignment Test from Selenium");
		numStudents.sendKeys("25");
		numQuestions.sendKeys("36");
		WebElement submit;
		submit = wait.until(ExpectedConditions.presenceOfElementLocated(new By.ByXPath("//*[@id=\"getting-started-buttons\"]/table/tbody/tr/td[2]/div/button")));
		submit.click();

	}

//	@After
//	public void tearDown() {
//		driver.quit();
//	}
}