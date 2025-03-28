# population-data-scraping
This is a simple example of web scraping using Python's Selenium WebDriver API. The purpose of this program is to extract the 2020 and 2021 populations as well as 2020-21 population % change for each U.S. state, the District of Columbia, Puerto Rico, and the U.S. overall (not including Puerto Rico). The program stores this data in the "2020-21 state population data" Excel workbook found in this folder by using the win32com Excel driver.

One test this program runs at the end is to ensure that the 2020 and 2021 U.S. overall populations are the sums of all state populations and the District of Columbia's population for each year.

Please note that you must download a version of ChromeDriver that works with your operating system and with your current Chrome version. Look for the right ChromeDriver version for you and download it here: https://googlechromelabs.github.io/chrome-for-testing/#stable
