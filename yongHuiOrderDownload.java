package com.ksj.connector.yonghui;

import java.awt.image.BufferedImage;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.text.DecimalFormat;
import java.text.FieldPosition;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.apache.http.HttpEntity;
import org.apache.http.client.CookieStore;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.protocol.ClientContext;
import org.apache.http.impl.client.BasicCookieStore;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.impl.cookie.BasicClientCookie;
import org.apache.http.protocol.BasicHttpContext;
import org.apache.http.protocol.HttpContext;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.openqa.selenium.By;
import org.openqa.selenium.Cookie;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.interactions.Actions;

import com.ksj.connector.yonghui.util.ConfigurationParser;
import com.ksj.connector.yonghui.util.MergeSalesFile;
import com.ksj.connector.yonghui.verification.VerificationUtil;

/**
 * 
 * @author guangjun.chen 201809
 *
 */
public class yongHuiOrderDownload {

	private static final String YONGHUISURL = "http://glzx.yonghui.cn:9000/newvssportal/login.html";
	private static String userName = "15981817108";
	private static String password = "sn258147";
	private static final String HOST = "glzx.yonghui.cn:9000";
	public static ConfigurationParser propertyUtil = null;
	public static final String SALESURL = "http://glzx.yonghui.cn/newvssportal/dataServer/mdqdxsrbb.html";
	private static String formattingString;
	private static DecimalFormat formatter;
	private static FieldPosition fPosition;
	private static StringBuffer buffer;

	public static void main(String[] args) {

		try {
			double start = System.currentTimeMillis();
			WebDriver driver = login();
			int retryTime = 0;
			int maxRetryTimes = 5;
			while (driver == null && retryTime <= maxRetryTimes) {
				retryTime++;
				Thread.sleep(5000);
				driver = login();
			}
			double end = System.currentTimeMillis();
			System.out.println("耗时" + (end - start) / 1000 + " s");
			if (driver != null) {
				System.out.println("开始永辉数据下载");
				// TODO
				driver.get(SALESURL);
				Thread.sleep(3000);
				// 循环食品公司
				driver.findElement(By.xpath("//i[@class='more-ico']")).click();
				
				List<WebElement> es = driver.findElements(By.xpath("//span[@class='num fl']"));
				int elementsSize = es.size();
				System.out.println("长度:"+elementsSize);
				Calendar c = Calendar.getInstance();
				SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				String dateStr2 = sdf2.format(c.getTime()).replace("-", "_").replace(":", "_").replace(" ", "_");
				List<String> txtFileList = new ArrayList<String>();
				String downloadDir = "E:/Data/YongHui/Sales/";
				for(int i =0;i<elementsSize;i++){
					
					System.out.println(es.get(i).getText());
					String venderCode = es.get(i).getText();
//					// 下载近一周sales数据
					for (int beforDays = 1; beforDays < 8; beforDays++) {
						Calendar c2 = Calendar.getInstance();
						c2.add(Calendar.DAY_OF_YEAR, -beforDays);
						SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy-MM-dd");
						String dateStr = sdf1.format(c2.getTime());
						String dateStr1 = dateStr.replaceAll("-", "");
						//点击查询
//						driver.findElement(By.xpath("//a[@class='btn search']")).click();
//						Thread.sleep(3000);
//						String dataEmpty = driver.findElement(By.id("code")).getAttribute("value").trim();
//						System.out.println("value:"+dataEmpty+"------"+ "".equals(dataEmpty));
//						if("".equals(dataEmpty)){//该食品公司暂无数据
//							continue;
//						}
						
						// http://glzx.yonghui.cn/newvss/vssService/MDXDXSRB_Service.do?method=getExportAllDetail&data={"appId":"GLZX_03","random":"","sign":"","data":"{\"startDate\":\"20180911\",\"endDate\":\"20180911\",\"venderCode\":\"20024912,\",\"shopId\":\"\"}"}
						
						String url = "http://glzx.yonghui.cn/newvss/vssService/MDXDXSRB_Service.do";
						url = url + "?method=getExportAllDetail";
						url = url + "&data={\"appId\":\"GLZX_03\",\"random\":\"\",\"sign\":\"\",";
						url = url + "\"data\":\"{\\\"startDate\\\":\\\""+dateStr1+"\\\",";
						url = url + "\\\"endDate\\\":\\\""+dateStr1+"\\\",";
						url = url + "\\\"venderCode\\\":\\\""+venderCode+",\\\",";
						url = url + "\\\"shopId\\\":\\\"\\\"}\"}";
						DefaultHttpClient client = null;
						Set<Cookie> cookies = driver.manage().getCookies();
						client = new DefaultHttpClient();
						CookieStore cs = new BasicCookieStore();
						for (Cookie ck : cookies) {
							BasicClientCookie bcc = new BasicClientCookie(ck.getName(),
									ck.getValue());
							bcc.setDomain(ck.getDomain());
							bcc.setPath(ck.getPath());
							cs.addCookie(bcc);
						}
						HttpContext localContext = new BasicHttpContext();
						localContext.setAttribute(ClientContext.COOKIE_STORE, cs);
						
//						HttpGet或 HttpPost都不能传包含 " 、“{"、"}"这样的参数，需要对特殊字符进行转义，把 "
//						 转成%22，把 { 转成%7b，把 } 转成%7d,把"\"转化为%5C
						url = url.replace(" ", "%20").replace("\"", "%22")
								.replace("{", "%7b").replace("}", "%7d").replace("\\",
										"%5C");
						HttpGet httpGet = new HttpGet(url);
						
						CloseableHttpResponse fileDownResponse =
								(CloseableHttpResponse) client.execute(httpGet,
										localContext);
						Thread.sleep(3000);
						HttpEntity entity = fileDownResponse.getEntity();
						
						File downDir = new File(downloadDir);
						if(!downDir.exists()){
							downDir.mkdirs();
						}
						String tempFileName ="KSJ.Yonghui_Sales."+venderCode+"."+dateStr1+".req"+ dateStr2 + ".xls";
						FileOutputStream fos = new
								java.io.FileOutputStream(downloadDir + File.separator +
										tempFileName);
						entity.writeTo(fos);
						fos.flush();
						fos.close();
						EntityUtils.consume(entity);
						fileDownResponse.close();
						
						System.out.println("源文件下载成功");
						
						String rootTempFileName = downloadDir + tempFileName;
						File tempFile = new File(rootTempFileName);
						String txtFileName = tempFileName.substring(0,
								tempFileName.lastIndexOf(".")) + ".txt";
						long fileLenth = tempFile.length();
						if(fileLenth>5000){
							txtFileList.add(txtFileName);
							String txtFileWithPath = downloadDir + txtFileName;
							
							//转化excel为txt
							convertXlsToTxtPOI(rootTempFileName, txtFileWithPath, '\t',
									0, 100, false, "");
						}
						
						tempFile.delete();
					}
				}
				
				MergeSalesFile.mergeFiles(downloadDir,txtFileList,"sales");


			} else {
				System.out.println("登录失败");
			}
		} catch (Exception e) {

			e.printStackTrace();
		}

	}

	public static boolean convertXlsToTxtPOI(String excelFile, String textFileName, char seperator,
			int numberOFLinesToSkip, int numberOFCellstokeep, boolean addHeader, String header) {
		boolean converted = false;
		try {

			int currentLine = 1;
			FileInputStream excelFileStream = new FileInputStream(new File(excelFile));

			Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(textFileName), "UTF-8"));

			org.apache.poi.ss.usermodel.Workbook workbook = null;
			if (excelFile.endsWith(".xls")) {
				workbook = WorkbookFactory.create(excelFileStream);
			} else {
				workbook = new XSSFWorkbook(excelFileStream);
			}
			org.apache.poi.ss.usermodel.Sheet datatypeSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = datatypeSheet.iterator();

			while (iterator.hasNext()) {
				String line = "";
				int cellNbr = 0;
				Row currentRow = iterator.next();
				Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {
					if (addHeader) {
						out.write(header);
						out.write("\r\n");
						out.flush();
						addHeader = false;
					}
					if (currentLine <= numberOFLinesToSkip) {
						break;
					}

					org.apache.poi.ss.usermodel.Cell currentCell = cellIterator.next();
					// currentCell.setCellStyle(null);
					// getCellTypeEnum shown as deprecated for version 3.15
					// getCellTypeEnum ill be renamed to getCellType starting
					// from version 4.0
					if (currentCell.getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING) {
						// System.out.print(currentCell.getStringCellValue() +
						// "--");
						if ("".equals(line)) {
							line = line + currentCell.getStringCellValue();
						} else {
							if (cellNbr < numberOFCellstokeep) {
								line = line + seperator + currentCell.getStringCellValue();
							}
						}
						// out.write(currentCell.getStringCellValue());
					} else if (currentCell.getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC) {
						// System.out.print(currentCell.getNumericCellValue() +
						// "--");
						// The StringBuffer is used to hold the result of
						// formatting
						// the number into a String. A new instance will be
						// required
						// each time.
						formattingString = "###0.#####";
						formatter = new java.text.DecimalFormat(formattingString);
						fPosition = new java.text.FieldPosition(0);
						buffer = new StringBuffer();

						// Recover the numeric value from the cell
						double value = currentCell.getNumericCellValue();

						// Format that number for display
						formatter.format(value, buffer, fPosition);

						// Not strictly necessary but I copy the result from the
						// StringBuffer into a String - leave this out for
						// performance
						// reasons in production code
						String resultString = buffer.toString();

						// Simply display the result to screen
						// System.out.println(resultString);
						if ("".equals(line)) {
							line = line + resultString;
						} else {
							if (cellNbr < numberOFCellstokeep) {
								line = line + seperator + resultString;
							}
						}
						// out.write("" + currentCell.getNumericCellValue());
					}

					// out.write(seperator);
					cellNbr++;
				}
				if (currentLine <= numberOFLinesToSkip) {
					currentLine = currentLine + 1;
					continue;
				}
				out.write(line);
				out.write("\r\n");
				out.flush();

			}
			out.close();
			converted = true;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return converted;
	}

	private static boolean DownloadSales(WebDriver driver) {

		return true;
	}

	public static WebDriver login() {

		propertyUtil = ConfigurationParser.getInstance();
		String gecoDriverPath = propertyUtil.getProperty("geco.driver.dir");
		System.setProperty("webdriver.gecko.driver", gecoDriverPath);


		WebDriver driver = null;
		FirefoxOptions options = new FirefoxOptions();
		 options.addArguments("--headless");
		driver = new FirefoxDriver(options);

		driver.manage().timeouts().implicitlyWait(600, TimeUnit.SECONDS);
//		driver.manage().timeouts().pageLoadTimeout(600, TimeUnit.SECONDS);// 设置页面加载等待时间
//		driver.manage().timeouts().setScriptTimeout(600, TimeUnit.SECONDS);
		driver.get(YONGHUISURL);
		System.out.println("log in to yongHui portal");

		try {
			WebElement ele = driver.findElement(By.xpath("//div[@id='qCode']/img"));

			// 获取验证码汉字坐标
			List<Point> pointList = VerificationUtil.downloadVerificationImage(driver,ele);

			driver.findElement(By.id("username")).clear();
			driver.findElement(By.id("username")).sendKeys(userName);
			System.out.println("Input user ID: " + userName);

			driver.findElement(By.id("password")).clear();
			driver.findElement(By.id("password")).sendKeys(password);
			System.out.println("Input PWD: " + password);

			Actions action = new Actions(driver);
			for (int i = 0; i < pointList.size(); i++) {
				action.moveToElement(ele, pointList.get(i).getX(), pointList.get(i).getY()).click().perform();

			}

			Thread.sleep(5000);

			String style = driver.findElement(By.xpath("//div[@class='qCode-ok']")).getAttribute("style");

			System.out.println("验证状态：" + style);

			// 点击登录按钮

			driver.findElement(By.xpath("//div[@class='btn-login']")).click();
			Thread.sleep(3000);

			String title = new String(driver.getTitle());
			if (title.contains("个人中心")) {

				System.out.println("login success");

			} else {
				System.out.println("login failed");
				driver.quit();
				driver = null;
			}

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("Connection Refused : " + e);
			driver.quit();
			driver = null;
		}
		return driver;
	}


}
