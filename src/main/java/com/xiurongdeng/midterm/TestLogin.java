package com.xiurongdeng.midterm;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class TestLogin {
  private final String pageURL;

  public TestLogin(String pageURL) {
    this.pageURL = pageURL;
  }

  public boolean login(String username, String password) {
    try {
      System.setProperty("webdriver.chrome.driver","src/main/resources/chromedriver");
      WebDriver driver = new ChromeDriver();
      driver.get(pageURL);
      driver.findElement(By.id("email")).sendKeys(username);
      driver.findElement(By.name("passwd")).sendKeys(password);
      driver.findElement(By.id("SubmitLogin")).submit();
      System.out.println("Login Done with Submit");
      return true;
    } catch (Exception e) {
      return false;
    }
  }
}
