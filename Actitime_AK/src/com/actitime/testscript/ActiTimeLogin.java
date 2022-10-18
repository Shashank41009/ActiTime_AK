package com.actitime.testscript;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.actitime.generic.FileLib;

public class ActiTimeLogin {
	
	static
	{
		System.setProperty("webdriver.chrome.driver", "./driver/chromedriver.exe");
	}

	public static void main(String[] args) throws IOException {
		FileLib f = new FileLib();
		WebDriver driver = new ChromeDriver();
		String url = f.getPropertyData("url");
		String un = f.getPropertyData("username");
		String pwd = f.getPropertyData("password");
		driver.get(url);
		driver.findElement(By.id("username")).sendKeys(un);
		driver.findElement(By.name("pwd")).sendKeys(pwd);
		driver.findElement(By.xpath("//div[text()='Login ']")).click();

	}

}
