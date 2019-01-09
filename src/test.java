import java.awt.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Objects;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;




public class test{
	
	public static void main(String args[]) throws IOException {
		HashMap <String, ArrayList> hashmap = new HashMap<String, ArrayList>();
		
		ArrayList<String> listOne = new ArrayList<String>();
		listOne.add("apple");
		listOne.add("aeroplane");
		
		ArrayList<String> listTwo = new ArrayList<String>();
		listTwo.add("bat");
		listTwo.add("Banana");
		
		ArrayList<String> listThree = new ArrayList<String>();
		listThree.add("Cat");
		listThree.add("car");
		
		hashmap.put("A", listOne);
		hashmap.put("B", listTwo);
		hashmap.put("C", listThree);
		
		System.out.println("hashmap : "+ hashmap);
		
	}
}
