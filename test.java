	package testcases;

	import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.aventstack.extentreports.reporter.configuration.ChartLocation;
import com.aventstack.extentreports.reporter.configuration.Theme;
import org.apache.log4j.Logger;


import java.util.Set;

	import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Proxy.ProxyType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.ITestResult;
import org.testng.SkipException;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

	import operation.UIoperations;
import operation.Utility;
import testcases.UploadComps;
import testcases.Download_comps;
import operation.Readobject;
import testcases.TradingComps;
import exportExcel.POIexcel;

	//@Listeners(testcases.Screenshot.class)

	public class Hybriddelete
	{   static int sweta;
	static int rowflag=0;
	int excelflag=0;
	FileInputStream inputStream;
		static WebDriver webdriver;
		String	FilepathforStatus="C:\\Users\\sweta.kumari\\Documents\\Repo\\CompsBuilder\\TestCases.xlsx";
		String getlist;
		public static ExtentHtmlReporter htmlReporter;
		public static ExtentReports extent;
		public static ExtentTest test;
		private Map<Object, String> sysprop;
		public static ExtentTest childTest;
	//	public String Testcasename2;
		//static Logger log = Logger.getLogger(HybridExecuteTest.class.getName());
	  public static Logger APP_LOGS = Logger.getLogger("devpinoyLogger");
		@Test(dataProvider="hybridData")
		public void testlogin(String testcasename,String keyword, String objectname,String objectType,String value,String browser) throws Exception
		{		
			if(keyword!=null)
		{
			if(testcasename!=null&&testcasename.length()!=0)
				{
				test = extent.createTest(testcasename,"Test Case: "+testcasename+ " Running on Browser: "+browser );
					APP_LOGS.info("\n\n" + "Started Executing test case-> " +testcasename+ "\n");
					if(browser.equals("IE"))
					{
						DesiredCapabilities capabilities = DesiredCapabilities.internetExplorer();
						capabilities.setCapability(CapabilityType.BROWSER_NAME, "IE");
						capabilities.setCapability(InternetExplorerDriver.
								INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,true);
						System.setProperty("webdriver.ie.driver", "C://Users//chaman.preet//Documents//C data//git//Comps_project//IEDriver//IEDriverServer.exe");
						webdriver=new InternetExplorerDriver(capabilities);
						webdriver.manage().window().maximize();
					}

					else if(browser.equals("Chrome"))
					{
						DesiredCapabilities capabilities = DesiredCapabilities.chrome();
						ChromeOptions options = new ChromeOptions();
						capabilities.setCapability(CapabilityType.BROWSER_NAME, "Chrome");
						 //capabilities.setCapability(ChromeOptions.CAPABILITY, options);
						System.setProperty("webdriver.chrome.driver", "C://Users//chaman.preet//Documents//C data//git//Comps_project//ChromeDriver//chromedriver.exe");
						webdriver=new ChromeDriver(capabilities);
						webdriver.manage().window().maximize();
					}
					
				else if(browser.equals("Firefox"))
				{
						FirefoxProfile fp = new FirefoxProfile();
						fp.setPreference("network.proxy.type", ProxyType.AUTODETECT.ordinal());
						System.setProperty("webdriver.gecko.driver", "C://Users//sweta.kumari//Downloads//selenium 3//geckodriver.exe");
						
						//download start from here
						
						
						fp.setPreference("browser.download.folderList", 2);
						fp.setPreference("browser.download.manager.showWhenStarting", false);
						fp.setPreference("browser.download.dir", "C:\\Users\\sweta.kumari\\Documents\\new\\");
						fp.setPreference("browser.helperApps.neverAsk.openFile",
								"text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml");
						fp.setPreference("browser.helperApps.neverAsk.saveToDisk",
				"text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml");
						fp.setPreference("browser.helperApps.neverAsk.saveToDisk", "application/zip;"+"application/x-zip;"+"application/x-zip-compressed;"+"application/octet-stream;"+"application/x-compress;"+"application/x-compressed;"+"multipart/x-zip;"+"application/x-unknown;"+"application/x-7z-compressed;"+".xlsx application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						fp.setPreference("browser.helperApps.alwaysAsk.force", false);
						fp.setPreference("browser.download.manager.alertOnEXEOpen", false);
						fp.setPreference("browser.download.manager.focusWhenStarting", false);
						fp.setPreference("wser.download.manager.useWindow", false);
						fp.setPreference("browser.download.manager.showAlertOnComplete", false);
						fp.setPreference("browser.download.manager.closeWhenDone", false);
					// download end here 
					//System.out.println("tst git  test command ");
						webdriver=new FirefoxDriver(fp);
						webdriver.manage().window().maximize();	
					}
				}
					
				 //browser end 
				Readobject robject=new Readobject();
				Properties allobjects=robject.getobjectrepository();
				UIoperations Uoperation=new UIoperations(webdriver);
				Uoperation.perform(allobjects, keyword, objectname, objectType, value);

				UploadComps objupload=new UploadComps(webdriver);
				Download_comps downobj=new Download_comps(webdriver);
				TradingComps tradeobj=new TradingComps(webdriver);
				downobj.download(allobjects, keyword, objectname, objectType, value);
				objupload.upload(allobjects, keyword, objectname, objectType, value);
				tradeobj.trade(allobjects, keyword, objectname, objectType, value);
	
		} //if end 
			else {
				throw new SkipException("test cases skipped ");
			}
			}
		
		//else if(runmode.equals("N"))
		//	{
			//	throw new SkipException("test cases skipped ");

	
	
		
		
		@SuppressWarnings("null")
		@DataProvider(name="hybridData")
		public Object[][] getDatafromDataprovider() throws Exception
		{  
		Object[][] object=new Object[1746][6];
		
		int x=0;
		//	Object[][] object=null;
		    POIexcel file=new POIexcel();
		    
			ArrayList<String> list=new ArrayList<String>();
			
			XSSFSheet sheet1=file.readexcel("C:\\Users\\sweta.kumari\\Documents\\Repo\\CompsBuilder", "TestCases.xlsx", "Test_cases");
			int rowcount1=sheet1.getLastRowNum()-sheet1.getFirstRowNum();
			System.out.println(rowcount1+"in first sheet ");
			int col_count1=sheet1.getRow(1).getPhysicalNumberOfCells();
			System.out.println("column of 1 sheet"+ col_count1);
			
			for(int i1=0;i1<rowcount1;i1++)
			{
				XSSFRow row1=sheet1.getRow(i1+1);
				for (int j1 = 0; j1 < row1.getLastCellNum(); j1++) {
					//Print excel data in console
					XSSFCell cell1=row1.getCell(j1);
					}
				//	object=new Object[rowcount1][col_count1];
				//	object[i1][j1] = cell1.toString();		
					//}
			//	System.out.println(object[0][2]);
			//file.getCellData("Test_cases","Runmode", 1);
				XSSFCell cell2=row1.getCell(2);
				
		if(cell2.toString().equals("Y"))
			{
				String Testcasename1 = row1.getCell(0).toString();
				System.out.println(Testcasename1);
				XSSFSheet sheet=file.readexcel("C:\\Users\\sweta.kumari\\Documents\\Repo\\CompsBuilder", "TestCases.xlsx", "test");
				int rowcount=sheet.getLastRowNum()-sheet.getFirstRowNum();
				System.out.println("row count is " +rowcount);
				int col_count=sheet.getRow(1).getPhysicalNumberOfCells();
				//object=new Object[rowcount][col_count];
				for(int i=0;i<rowcount;i++)
				{
					XSSFRow row=sheet.getRow(i+1);
					for (int j = 0; j < row.getLastCellNum(); j++) 
					{
						//Print excel data in console
						XSSFCell cell=row.getCell(j);
						//object[i][j] = cell.toString();						
						//System.out.println("values are"+" " +object[i][j]);
						
					}
					String Testcasename2 = row.getCell(0).toString();
			//	System.out.println("value of TC2 is " +Testcasename2);	
					if(Testcasename1.equalsIgnoreCase(Testcasename2))
					{System.out.println(row.getRowNum());			
					 x=row.getRowNum();
					System.out.println("swetacount "+x);
						System.out.println("Testcase name matches " +Testcasename2);
						do{
													//	int Row_blankcount = row.getCell(0).CELL_TYPE_STRING;
							//System.out.println(Row_blankcount);
							System.out.println("same test case");
												
							XSSFRow row3=sheet.getRow(x);
					
						for (int z = 0; z < row3.getLastCellNum(); z++) 
						{
							//Print excel data in console
							XSSFCell cell=row3.getCell(z);
							object[x][z] = cell.toString();
							//String c = cell.toString();
							//object2 = object[x][z];
						//	System.out.println("values are "  +object[x][z]);
							//list.add(c);
						//	dp.add(new Object[] {c});
						}x++;
					} 
						//while(row.getCell(0).getStringCellValue().isEmpty() || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK || row.getCell(0).toString()==null);
						while(sheet.getRow(x).getCell(0).getStringCellValue().isEmpty() || sheet.getRow(x).getCell(0).getCellType() == Cell.CELL_TYPE_BLANK || sheet.getRow(x).getCell(0).toString()==null);
						sweta=x;
					}
					else {System.out.println("Next testcase started"  );}
				}
		
						
		
			System.out.println("object value are "+object);
					
			//object[x][z] = cell.toString();
			
			}	
		else if(cell2.toString().equals("N"))
		{
		
		System.out.println("Runmode is N");}

	}
		
		//	Object[][] object1=object1[sweta][6];
			//Object[][] object1=new Object[sweta][6];
			System.out.println("value out of for loop"+ x);
			return object;
		}
		
	//}
		@AfterMethod
		public void screenshot(ITestResult result) throws IOException
		{
			try 
			{ // extent.endTest(test);
				String methodname1=result.getName().toString().trim();
				String methodname= result.getName()+ "-" + Arrays.toString(result.getParameters());
				String report=methodname.substring(0,40);
				String scrnshotname=methodname.substring(12,30 );
				if(ITestResult.SUCCESS==result.getStatus() ||  ITestResult.FAILURE==result.getStatus())
				{
					APP_LOGS.info("Test step is -> " +result.getName()+ "-" + Arrays.toString(result.getParameters()));

				}

				if(ITestResult.SUCCESS==result.getStatus())
				{
					APP_LOGS.info("PASS");
					FileInputStream inputStream;
					try {
						inputStream = new FileInputStream(FilepathforStatus);
						Workbook myworkbbok = null;
						myworkbbok = new XSSFWorkbook(inputStream);
						Sheet sheet = myworkbbok.getSheet("Comps");
						XSSFRow row = (XSSFRow) sheet.getRow(excelflag+1);
						XSSFCell cell1 = row.getCell(7);
						cell1.setCellValue("pass");
						inputStream.close();
						FileOutputStream outputStream = new FileOutputStream(FilepathforStatus);
						myworkbbok.write(outputStream);
						outputStream.close();
						excelflag++;

					} catch (Exception e1) {
						e1.printStackTrace();
					}
					//System.out.println("Screenshot taken for pass test cases  ");
					Utility.capturescreenshot(webdriver, scrnshotname,"C:\\Users\\sweta.kumari\\Documents\\Repo\\CompsBuilder\\ScreenShotForPassTestCases\\");
					test.pass(MarkupHelper.createLabel(report+" Test Step PASSED", ExtentColor.GREEN));

				} 


				else if (ITestResult.FAILURE==result.getStatus())	
				{	
					Throwable cause = result.getThrowable();
					if (null != cause) {
						APP_LOGS.error(" **FAIL - " +cause.getMessage());}

					try {
						inputStream = new FileInputStream(FilepathforStatus);
						Workbook myworkbbok = null;
						myworkbbok = new XSSFWorkbook(inputStream);
						Sheet sheet = myworkbbok.getSheet("Comps");
						XSSFRow row = (XSSFRow) sheet.getRow(excelflag+1);
						XSSFCell cell1 = row.getCell(7);
						cell1.setCellValue("fail");
						inputStream.close();
						FileOutputStream outputStream = new FileOutputStream(FilepathforStatus);
						myworkbbok.write(outputStream);
						outputStream.close();
						excelflag++;
					} 
					catch (Exception e1)
					{
						e1.printStackTrace();
					}
					Utility.capturescreenshot(webdriver, scrnshotname,"C:\\Users\\sweta.kumari\\Documents\\Repo\\CompsBuilder\\ScreenShotForFailTestCases\\");
					test.fail(MarkupHelper.createLabel(report+" Test Step failed", ExtentColor.RED));

					String screenShotPath = Utility.capture(webdriver, scrnshotname);
					test.addScreenCaptureFromPath(screenShotPath);
				}
				else if (ITestResult.SKIP==result.getStatus())	
				{
					try {
						inputStream = new FileInputStream(FilepathforStatus);
						Workbook myworkbbok = null;
						myworkbbok = new XSSFWorkbook(inputStream);
						Sheet sheet = myworkbbok.getSheet("Comps");
						XSSFRow row = (XSSFRow) sheet.getRow(excelflag+1);
						XSSFCell cell1 = row.getCell(7);
						cell1.setCellValue("skip");
						inputStream.close();
						FileOutputStream outputStream = new FileOutputStream(FilepathforStatus);
						myworkbbok.write(outputStream);
						outputStream.close();
						excelflag++;
					} catch (Exception e) {
						e.printStackTrace();			}

					test.skip(MarkupHelper.createLabel(report+" Test Step SKIPPED", ExtentColor.ORANGE));
				}
			}
			catch (Exception e)
			{

				System.out.println("Exception while taking screenshot "+e.getMessage());
			} 


		}
		@BeforeSuite
		public void beforesuite() throws UnknownHostException
		{System.out.println("test");
			String username = System.getProperty("user.name");
			String OS=System.getProperty("os.name");
			String Hostname=InetAddress.getLocalHost().getHostName();
			htmlReporter = new ExtentHtmlReporter(System.getProperty("user.dir") +"/test-output/SwetaReport.html");
			extent = new ExtentReports();
			extent.attachReporter(htmlReporter);
			htmlReporter.loadConfig((System.getProperty("user.dir") +"/extent-config.xml"));
			extent.setSystemInfo("OS", OS);
			extent.setSystemInfo("Host Name",Hostname);
			extent.setSystemInfo("Environment", "QA");
			extent.setSystemInfo("User Name", username);
			htmlReporter.config().setChartVisibilityOnOpen(true);
			htmlReporter.config().setDocumentTitle("CompsBuilder Automation Report ");
			htmlReporter.config().setReportName("Comps Regression Testing Report ");
			htmlReporter.config().setTestViewChartLocation(ChartLocation.TOP);
			htmlReporter.config().setTheme(Theme.STANDARD);
		}
		@AfterSuite
		public void aftersuite()
		{
			extent.flush();
		}

	
	}
