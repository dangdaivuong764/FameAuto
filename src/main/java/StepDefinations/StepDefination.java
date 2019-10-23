package StepDefinations;

import java.io.IOException;

import org.testng.annotations.Test;

import OperationHelper.OperationHelper;
import cucumber.api.DataTable;
import cucumber.api.java.After;
import cucumber.api.java.Before;
import cucumber.api.java.en.Then;

public class StepDefination {
	OperationHelper op = new OperationHelper();
	
	@Before
	public static void browser() throws IOException, InterruptedException {
		OperationHelper.OpenBrowser();
	}
	@Then("^I want to run test case form file excel with \"([^\"]*)\" and \"([^\"]*)\" write result to round test \"([^\"]*)\"$")
	public void i_want_to_run_test_case_form_file_excel_with_and(String fileName, String sheetName, String numberRound, DataTable data)
			throws Throwable {
		op.CheckUsercaseOnExcel(fileName, sheetName, numberRound, data);
	}
	@After
	public void sterDown() {
		OperationHelper.destroy();
	}
	/*
	@Test
	public void testing() throws Exception {
		op.checkFileExcel("TestingResult.xlsx", "Test_Cases");
	}
	@Given("^Open the Firefox and launch the application$")
	public void open_the_Firefox_and_launch_the_application() throws Throwable {
		String openBrowser = ReadProperties.getElementsAndUrls("url");
		String getLocalName = ReadProperties.getElementsAndUrls("localname");
		OperationHelper op = new OperationHelper();
		op.getLocalName(getLocalName, openBrowser);
	}
	@Given("^I want to write a step with ([^\"]*) and ([^\\\"]*)$")
	public void i_want_to_write_a_step_with_name_and_test(String username, String password) throws Throwable {
		System.out.println("User Name is  " + username);
		System.out.println("Password is  " + password);
		op.TestFaceBook(username, password);
	}
	 */
}
