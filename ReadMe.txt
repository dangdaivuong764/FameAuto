1/ Fix cannot install TestNG: https://www.techbeamers.com/install-testng-in-eclipse-ide/
2/ Xu ly dynamic element
+ Absolut xpath: web_element_name=html/body/div[30]/div[2]/div[2]/div/div/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/em
+ Identify Element by starting Text: //button[starts-with(@id, 'Submit-')]
+ Identify Element by containing Text: //input[contains(@class, 'suggest')].
+ Identify Element by Index: driver.findElements(By.xpath(“//*submit”)).get(0).click();
+ Identify Element with reference to a closest stable element : //span1/../following-sibling::div//button1

3/ Comapre basic two image:
https://github.com/assertthat/selenium-shutterbug/wiki/Examples-of-usage

4/ 	if (error.length() > 0) {
					cell = sheet.getRow(row).createCell(resultInex);
					cell.setCellValue(error);
					fileOut = new FileOutputStream(strFileName);
					workbook.write(fileOut);
					//fileOut.flush(); cannot write data to excel when it appear
					fileOut.close();
					System.out.println("ERROR " + error);
					//error = "";
					break;
				}
