package com.basicTest;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class ParallelTestingDemo2 {
	
	@Test //(threadPoolSize =3, invocationCount = 3)
	public void Test1() throws InterruptedException{
		
		System.out.println("I am inside Test1 |" + Thread.currentThread().getId());
		
		System.setProperty("webdriver.gecko.driver","C:\\Users\\prabo\\eclipse-workspace\\W_TestNGG\\driverFiles\\geckodriver.exe");
	    WebDriver driver= new FirefoxDriver();
	    driver.get("https://www.google.com/");
	    Thread.sleep(3000);
	  //  driver.close();
		
	}
	@Test
	public void Test2() throws InterruptedException{

		System.out.println("I am inside Test2  |" + Thread.currentThread().getId());		

		System.setProperty("webdriver.gecko.driver","C:\\Users\\prabo\\eclipse-workspace\\W_TestNGG\\driverFiles\\geckodriver.exe");
	    WebDriver driver= new FirefoxDriver();
	    driver.get("https://www.walmart.com/");
	    Thread.sleep(3000);
	 //   driver.close();
		
	}
	
	@Test
	public void Test3() throws InterruptedException{
		
		System.out.println("I am inside Test3 |" + Thread.currentThread().getId());	
	
	}
}
