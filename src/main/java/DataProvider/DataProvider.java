package DataProvider;

import java.io.File;
import java.io.IOException;

public class DataProvider {
	/*get path from Folder
	 * param : fileName*/
	public String getPropertiesFile(String nameFile) throws IOException {
		String path = System.getProperty("user.dir") + File.separator + "src" 
					  + File.separator + "test" +  File.separator +"properties" +  File.separator + nameFile;
		return path;
	}
	public String getExcelfile(String nameFile) throws IOException {
		String path = System.getProperty("user.dir") + File.separator + "src" 
					  + File.separator + "test" +  File.separator +"fileExcel" +  File.separator + nameFile;
		return path;
	}
	public String getStoreScreenShoot(String nameStore) {
		String pathScreenShoot = System.getProperty("user.dir") + File.separator + "src" 
				  + File.separator + "test" +  File.separator +"screenShoot" +  File.separator + nameStore;
		return pathScreenShoot;
	}
}
