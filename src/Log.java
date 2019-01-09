import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.ImageIcon;
import javax.swing.JToolBar;
import javax.swing.JButton;
import javax.swing.JTextField;
import javax.imageio.ImageIO;
import javax.sound.sampled.AudioSystem;
import javax.sound.sampled.Clip;
import javax.sound.sampled.LineUnavailableException;
import javax.sound.sampled.UnsupportedAudioFileException;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.awt.event.ActionEvent;
import javax.swing.JTextArea;
import javax.swing.border.LineBorder;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Image;

import javax.swing.JScrollPane;
import javax.swing.WindowConstants;

public class Log extends JFrame {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

														 //must declare objects before even main method to use in all of the class, must be static
	static LinkedHashMap <Integer, Students> objMap = new LinkedHashMap <Integer, Students>();
	static LinkedHashMap <Integer, String> studentMap = new LinkedHashMap<Integer, String>();
	
	private JPanel contentPane;			//create content pane for the GUI
	private JTextField mLogin;			//create a text field for the GUI
	public static JTextArea txtArea = new JTextArea(); 		//create a text area for live logging

//**********************************************************************************Main Method********************************************************************************************************************************
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {							//create a runnable method, starts the main logging system
				try {
					
					
					Log frame = new Log();				//runs the log() method 
					frame.setVisible(true);				//sets the main window visible
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}); 
	}

	/**
	 * Create the frame.
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 */
	public Log() throws InvalidFormatException, IOException {
		addWindowListener(new java.awt.event.WindowAdapter(){		//makes the window listen when you try to close it, intercepts the closing and askes if you really want to close
			public void windowClosing(java.awt.event.WindowEvent windowEvent){
		        if (JOptionPane.showConfirmDialog(contentPane ,"Are you sure you want to stop logging time? \r\n (all students logged in will be logged out)", "Really Close?", 
		        		JOptionPane.YES_NO_OPTION,
		                JOptionPane.QUESTION_MESSAGE) == JOptionPane.YES_OPTION){
		                try {
							shutdownCheck();		//calls the shut down check method
						} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
							e.printStackTrace();
						}
		            }
			}

		});
		
        setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);			//keeps the program from closing if you select no in the shutdown dialog box

        final JPanel panel = new JPanel();  					//creates the panel for the pictures in the program
        panel.setPreferredSize(new Dimension(1620, 3000));		//sets the size of the panel for the pictures, if pictures are being cut off at the bottom increase the height value

        final JScrollPane scroll = new JScrollPane(panel);		//creates a scrolling pane to put the above panel into to allow scrolling
        scroll.setBounds(0, 116, 2044, 1571);					//sets the location and size of the scrolling pane
        scroll.getVerticalScrollBar().setUnitIncrement(16);		//sets the scroll speed of the vertical scroll bar
        scroll.getVerticalScrollBar().setPreferredSize(new Dimension(30, 0));		//sets the size of the vertical scroll bar
        panel.setLayout(null);
        
        getContentPane().setLayout(null);			//gets the base content pane of the program and sets the layout style to null
        getContentPane().add(scroll);				//adds the scroll pane to the base content pane
        
        JLabel lblTimeLoggingSystem = new JLabel("Time Logging System (6164)");			//creates a JLabel for the title of the program
        lblTimeLoggingSystem.setBounds(913, 0, 1076, 118);								//sets the location and size of the JLabel
        lblTimeLoggingSystem.setFont(new Font("Magneto", Font.PLAIN, 65));				//sets the font of the JLabel
        getContentPane().add(lblTimeLoggingSystem);										//adds the JLabel to the base content pane
        
        
		//adds new entries into the sutdentMap array, to add new students enter them into this list .put(studentID, student name)
        studentMap.put(2809, "Joanna Seymour");
		studentMap.put(1, "Tyler Strickler");
		studentMap.put(3, "Paul Rael");
		studentMap.put(4, "Gail Strickler");
		studentMap.put(5, "John Latusek");
		studentMap.put(6, "Todd Graper");
		studentMap.put(7,"Molly Burkle");
		studentMap.put(8,"Matt Miller");
		studentMap.put(9,"Don McCallum");
        studentMap.put(2242, "Cassidy Borwig");
		studentMap.put(1856, "Dylan Brewer");
		studentMap.put(2390, "Braden Brown");
		studentMap.put(2456, "Taylor Brown");
		studentMap.put(2391, "Zachary Clark");
		studentMap.put(2288, "Tristin Cleveland");
		studentMap.put(1881, "Maya Despenas");
		studentMap.put(2394, "Jayne Eilderts");
		studentMap.put(2035, "Victoria Fernandez");
		studentMap.put(2129, "Colton Glick");
		studentMap.put(2401, "Katryna Hauser");
		studentMap.put(3348, "Adrian Hinz");
		studentMap.put(2573, "Clarissa Lentfer");
		studentMap.put(2109, "Sydney Hoffmann");
		studentMap.put(2406, "Emilee Junker");
		studentMap.put(2447, "Tyler Laube");
		studentMap.put(2583, "Alana Ledtje");
		studentMap.put(2086, "Sawyer Loger");
		studentMap.put(2291, "Jon McCallum");
		studentMap.put(2417, "Madyson Mrzlak");
		studentMap.put(1887, "Adriana Murphy");
		studentMap.put(2418, "Malea Neuroth");
		studentMap.put(2005, "Lauren Vanderlind");
		studentMap.put(2425, "Michaela Wagner");
		studentMap.put(2755, "Katrina Miller");
		studentMap.put(3345, "Brandon Newcomb");
		studentMap.put(3070, "Kiana Gardner");
		studentMap.put(3503, "Taylor Stanbrough");
		studentMap.put(2320, "Blain Stewart");
		studentMap.put(2444, "Delphia McCallum");
		
				/* Setup each student object and image */
		
		
		int tileX = 10; 				//tileX and tileY are the locations of each images, it is set to 10,10 for the first image then increased later in the loop 
		int tileY = 10;
		
		int tileBuffX = 30;				//the buffer between images 
		int tileBuffY = 50;
		
		int imageWidth = 131*2;			//sets the image size (if size is increased spacing will also have to be increased)
        int imageHeight = 174*2;		//The raw image can be what ever size, the program will scale it to this size
        
		LinkedHashMap <Integer, String> middleMap = new LinkedHashMap<Integer, String>();		//to solve issue where objects would be taken out of the studentMap 
		middleMap.putAll(studentMap);															//this middleMap allows me to iterate through the array without harming the original

	    Iterator it = middleMap.entrySet().iterator();			/* create a 'Students' object for each entry in the middleMap */
	    while (it.hasNext()) {
	        Map.Entry pair = (Map.Entry)it.next();  
	        final Students student = new Students();			//creates the base object
	        student.studentID = (int) pair.getKey();			//sets the studentID based on the key from the array
	        student.name = (String) pair.getValue();			//sets the name based on the value from the array
	        student.state = false;								//sets the state to false by default
	        student.image = new JLabel(student.name);			//creates a new JLabel for each student and sets the text to the students name
	        
	        							
	        			/* image size may vary but a ratio equal to 175x232 is suggested, supports only GIF, PNG, JPEG, BMP, and WBMP */
	        			/* The image name must be exactly the same (case sensitive) as the name of the student in the studentMap */
	        
	        if (Log.class.getResource("/images/"+student.name+ ".png")!=null){									//If there is an image with the name of the student run this code
	        	URL imageLocation = this.getClass().getResource("/images/"+student.name+".png");				//get the URL for the image
	        	BufferedImage intiImage =ImageIO.read(imageLocation);											//save the image into a buffer
	        	Image scaledImage  = intiImage.getScaledInstance(imageWidth, imageHeight, Image.SCALE_FAST);	//scale the image to the desired width and height
	        	student.image.setIcon(new ImageIcon(scaledImage));												//set the scaledImage to the current student's image
	        }
	        else{																								//If there is NOT an image with the name of the student run this code
	        	URL imageLocation = this.getClass().getResource("/images/default.png");							//saves the URL for the default image
	        	BufferedImage intiImage =ImageIO.read(imageLocation);											//saves the image into a buffer
	        	Image scaledImage  = intiImage.getScaledInstance(imageWidth, imageHeight, Image.SCALE_FAST);	//scales the image to the desired width and height
	        	student.image.setIcon(new ImageIcon(scaledImage));												//sets the scaledImage to the current student's Image
	        	
	        }
	        
	        student.image.setName(student.name);								//sets the image name to the same as the students
		    student.image.setToolTipText(student.name);							//sets the tool tip of the image to the student's name
		    student.image.setEnabled(student.state);							//sets the enabled to false (the same as the student.state)
		    student.image.setBounds(tileX,tileY,imageWidth,imageHeight);		//sets the location and size of the image
		    
			    if (tileX < 1500){								//if tileX is less than 1500 add the next image to the right
			    	tileX = tileX + imageWidth + tileBuffX;		//sets the location for the next image to the current location plus the width plus the buffer
			    }
			    
			    else{											//if tileX is larger than 1500 set tileX to 10 and set tileY to the current location plus the height plus the buffer
			    	tileX = 10;
			    	tileY = tileY + imageHeight+tileBuffY;
			    }
			objMap.put(student.studentID, student);						//adds each student image to an object map to more easily reference the student object through the student ID
		    panel.add(student.image);									//finally adds the image to the panel for the user to see
		    student.image.addMouseListener(new MouseAdapter(){			//adds a listener event to the image to listen for a mouse click on the image 
		    	@Override
		    	public void mouseClicked(MouseEvent arg0){						//if mouse clicks image do the following
		    		if(student.image.isEnabled()){								//if image is enabled / student.state is true
		    			try{
		    				writeFile(student.name, false, student.studentID);	//calls the writeFile function to log the time logged out
		    				student.image.setEnabled(false);					//set the image to be disabled	
		    				student.state=false;								//sets the student.state to false
		    			} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
		    				e.printStackTrace();
		    			}
		    		}
		    	}
		    	
		    });
	        it.remove(); // avoids a ConcurrentModificationException
	    }

		
		//creates the toolbar and adds it to the content pane
        JToolBar toolBar = new JToolBar();				//create the toolbar element for the top of the window
        toolBar.setFloatable(false);					//sets the floatability of the toolbar to false
        toolBar.setBounds(0, 0, 631, 95);				//sets the location and size of the toolbar
        getContentPane().add(toolBar);					//adds the toolbar to the base content pane
        setSize(2747, 1758);							//sets the size
		
		
		//creates text field for logging in and adds it 
		mLogin = new JTextField();									//creates a text field to enter the user ID 
		mLogin.setFont(new Font("Tahoma", Font.PLAIN, 57));			//sets the font of the text field
		mLogin.setText("Manual Login");								//sets the defualt text of the text field
		toolBar.add(mLogin);										//adds the text field to the tool bar
		mLogin.setColumns(10);				
		txtArea.setLineWrap(true);									//sets line wrap to true
		
		//creates the text area for the live logging and adds it
		txtArea.setEditable(false);										//sets the text area to non-editable
		txtArea.setFont(new Font("Courier New", Font.PLAIN, 30));		//sets the fond of the text area
		txtArea.setText("CONSOLE LOG:\r\n");							//sets the default text to write to the text area
		txtArea.setBounds(2054, 116, 667, 1571);						//sets the location and size of the text area
		getContentPane().add(txtArea);									//adds the text area to the default content pane
		txtArea.setBorder(new LineBorder(new Color(0, 0, 0), 3, true));	//sets the border style of the text area
	
		final JButton btnSubmit = new JButton("Submit");		//creates submit button
		btnSubmit.setFont(new Font("Tahoma", Font.PLAIN, 57));	//sets the font of the submit button
		btnSubmit.addActionListener(new ActionListener() {		//listens
			public void actionPerformed(ActionEvent arg0) {
				int mLoginNumber = Integer.parseInt(mLogin.getText());		//turns any text in the text field into an integer and saves it to the variable 'mLoginNumber'
				if(studentMap.containsKey(mLoginNumber)){					//if the hashmap from Students class contains that number in one of the keys it continues else it displays an error
					String username = studentMap.get(mLoginNumber);	    	//sets the username variable to the value of the ID in the studentMap	
					
/*set state */	Students currentS= null;												//creates a new "Students" object and sets it to null
					for(Map.Entry entry : objMap.entrySet()){							
						if((int) entry.getKey()==mLoginNumber){							//if the current entry is the loginNumber that was provided run the following
							currentS=(Students) entry.getValue();						//set the currentS to the value of the entry from the objMap
							currentS.state = !currentS.state;							//change the state to the opposite of what it is currently
							currentS.image.setEnabled(!currentS.image.isEnabled());		//set the image to enabled
							break;
						}
					}
					
					
					if (!currentS.state){											//if after the change the state is false run the following block (logging out)
							try {
								writeFile(username, false, mLoginNumber);			//calls the 'writeFile' method, passes the following information 'username' False and 'mLoginNumber' 
							} catch (FileNotFoundException e) {						//'writeFile(Name of the user, Are they logging in? , what is the users ID);
								System.out.println("File not found");
								e.printStackTrace();
							} catch (IOException e) {
								System.out.println("IO Exception");
								e.printStackTrace();
							} catch (EncryptedDocumentException e) {
								System.out.println("Encrypted Document Exception");
								e.printStackTrace();
							} catch (InvalidFormatException e) {
								System.out.println("Invalid Format Exception");
								e.printStackTrace();
							}
						}
					else{																//If after the change the state is true run the following block (Logging in)
					try {																	
						writeFile(username, true, mLoginNumber);						//calls 'writeFile'
					} catch (IOException e) {											//'writeFile(Name of the user, Are they logging in? , what is the users ID);
						System.out.println("IO Exception");
						e.printStackTrace();
					} catch (EncryptedDocumentException e) {
						System.out.println("Encrypted Document Exception");
						e.printStackTrace();
					} catch (InvalidFormatException e) {
						System.out.println("Invalid Format Exception");
						e.printStackTrace();
					}
					}
				}
				else{
					JOptionPane.showMessageDialog(null, "Theres no user by the code : "+ mLogin.getText());		//message that is displayed if no user id is found
				}
				mLogin.setText("");			//sets the text in the login field to nothing
				mLogin.requestFocus();		//sets the login field to the main focus for entering the next student ID
			}
		});
		
		toolBar.add(btnSubmit);			//finally add the submit button to the toolbar
		toolBar.getRootPane().setDefaultButton(btnSubmit);
		
	}
	
	
	public static void writeFile (String user, boolean login, int mLoginNumber) throws IOException, EncryptedDocumentException, InvalidFormatException{			//new writeFile method (runs everytime someone logins or out)
		String workingDir = System.getProperty("user.dir") + "\\Activity Log.xls";			//gets the location of where the excel log file should/will be, (is a string)
		Path fileCheck = Paths.get(workingDir);												//converts the String variable workingDir to a path variable
		if(!Files.exists(fileCheck)){														//checks if the excel file exsits, If it does not exsist run the following block of code
			String fileName = "Activity Log.xls";											//sets the file name 
			HSSFWorkbook workbook = new HSSFWorkbook();										//creates a new excel workbook to write in
			Sheet sheet1 = workbook.createSheet("sheet1");									//creates a new sheet called sheet1 in the workbook
				
			writeCell(sheet1,0,0,"Date v | Name >");							//calls the write cell method, writeCell(name of the sheet, x coordinate , y coordinate , text to be written);
			writeCell(sheet1,0,1,"Total Time (hrs.) >");
			writeCell(sheet1,0,2,getDate());									//writeCell(name of the sheet, x coordinate , y coordinate , text to be written);
			
			//get a set of the entries
			Set set = studentMap.entrySet();			//create a set with all the entries of the hashmap in the students class
					
			//get an iterator
			Iterator i = set.iterator();				//create an iterator (still don't fully understand this)
			int x=1;									//create an x variable and set it to 1
			
			while(i.hasNext()){									//Start a while loop, while there are more people to look through keep running this block of code
				Map.Entry me = (Map.Entry)i.next();				//create a map entry (still don't fully understand this)
				String userLoop = (String) me.getValue();		//gets the name of the current user it is looking at. saves it to variable 'userLoop'
				writeCell(sheet1,x,0,userLoop);					//calls writeCell method, writeCell(name of the sheet, x coordinate , y coordinate , text to be written);
				x++;											//adds 1 to x, makes the cell that is to be written to move, i controls what person it is looking at in the hashmap
			}	
			FileOutputStream fileOut = new FileOutputStream(fileName);		//set up a new output stream to link java and the excel file
			workbook.write(fileOut);										//after the loop ends it sends all that data to the 'fileOut' output stream to be written to the excel file
			workbook.close();												//closes the workbook to finish
			fileOut.close();												//closes the fileOut stream
			
			writeFile(user,login,mLoginNumber); 			//runs the entire method again, this makes sure to log the person who made the program create the excel file
		}
		else{												//If the file does exist run this code													
			
			String fileName = "Activity Log.xls";					//sets the file name
			String currentTime = getTime();							//runs the 'getTime' method and saves its return value to 'currentTime'
			InputStream inp = new FileInputStream(fileName);		//Creates an input stream to read whats in the file
			Workbook wb = WorkbookFactory.create(inp);				//Create a work book with the workbookFactory method
			Sheet sheet = wb.getSheet("sheet1");					//create a sheet by getting the sheet called sheet1 in the workbook
			int colNum = sheet.getRow(0).getLastCellNum();			//creates an int colNum and gets the last cell in the first row
			int cellX = findX(user);								//runs the 'findX' method and saves its return vales to 'cellX' , findX(name of user);
				if(cellX==0){										//if the findX returns a value of 0 (cant find the user) run the following
					writeCell(sheet,colNum,0,user);					// write the users name to the left most open cell in the first row (colNum)
					cellX=colNum;									// sets cellX to colNum
				}
				
			int cellY = findY(cellX);						//runs the 'findY' method and saves its return value to 'cellY' , findY(x value);
			
			Row yCord = sheet.getRow(cellY);				// create a Row variable called 'yCord' from the 'cellY' variable
			Cell targetCell = yCord.getCell(cellX);			//create a cell variable called 'targetCell' by using yCord and getting the cell value from 'cellX'
				
			/* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- */
				
				if(login){																//if login is true (you are logging in) run this block of code
					writeCell(sheet,cellX,cellY,"Log in @"+currentTime);				//calls the 'writeCell' method, writeCell(name of the sheet, x coordinate , y coordinate , text to be written);
					URL loginSound = Log.class.getResource("/sounds/login.wav");		//gets the login sound effect and saves it to the URL variable 'loginSound'
					playSound(loginSound);												//calls the 'playSound' method
					
					/*writes update to live logging window*/
					
					if (txtArea.getLineCount()>= 28){									//if the line count is greater than 28 erase what is currently in the window
						txtArea.setText(null);
					}
					txtArea.append(user + " logged in at:" + currentTime +"\r\n");		//does not write to excel file, writes to the live logging in the GUI
					
				
				}
				else{															// if login is false (you are logging out) run this block of code
					writeCell(sheet,cellX,cellY,"Log out @"+currentTime);		//calls 'writeCell', writeCell(name of the sheet, x coordinate , y coordinate , text to be written);
					
					String cellString = cellToString(targetCell);				//converts the target cell to a string
					
					Row totalTime = sheet.getRow(1);																	//gets the row for total times
					Cell timeTarget = totalTime.getCell(cellX);															//gets the cell for the total time of the correct user
					int loginLocation = cellString.lastIndexOf("Log in @");												//get the last index of "Log in @" in the loginLocation cell
					float newTime =hourComp(cellString.substring(loginLocation + 9, loginLocation+14),currentTime);		//calls the hourComp function to find new total, hourComp(login time (ex: 13,56), logout time)
					String newTimeString = String.valueOf(newTime);

					if (timeTarget==null){
						setCell(sheet,cellX,1,newTimeString);
					}
					else{
						String currentTotalString = timeTarget.getStringCellValue();
						float currentTotal = Float.parseFloat(currentTotalString);
						float newTotal = currentTotal+newTime;
						newTimeString = String.valueOf(newTotal);
						setCell(sheet,cellX,1,newTimeString);
					}
					
					
					URL logoutSound = Log.class.getResource("/sounds/logout.wav");			//gets the logout sound
					playSound(logoutSound);													//plays the logout sound
					if (txtArea.getLineCount()>= 28){
						txtArea.setText(null);
					}
					txtArea.append(user + " logged out at:" + currentTime +"\r\n");			//write to the live logging in the GUI
					txtArea.append(user + "'s new total is: " + newTimeString + "Hrs." +"\r\n");	
				}
				FileOutputStream fileOut = new FileOutputStream(fileName);					//create the fileoutput stream to write to the excel file
				wb.write(fileOut);						//write to the file
				fileOut.close();						//close the file stream
		} 
	}
	public static String getDate(){								//gets the current date on the system this program is running from
		DateFormat df = new SimpleDateFormat(" MM-dd-yy ");
		Date dateobj = new Date();
		return  df.format(dateobj);
	}
	public static String getTime(){								//gets the current time on the system this program is running from
		DateFormat df = new SimpleDateFormat(" HH,mm ");		//cant use : so replaced with ,
		Date dateobj = new Date();
		return  df.format(dateobj);
	}
	public static int findX(String username) throws IOException, EncryptedDocumentException, InvalidFormatException{		//find the x value by searching for the users name
		int result=0;												//creates result variable, will use this at the end of the method
		String fileName = "Activity Log.xls";						//sets the file name
		InputStream inp = new FileInputStream(fileName);			//creates a input stream to read the file
		Workbook wb = WorkbookFactory.create(inp);					//create a workbook from the input stream
		Sheet ws = wb.getSheet("sheet1");							//get the sheet called sheet1 in the workbook
			
			
			int colNum = ws.getRow(0).getLastCellNum();		// colNum is the max x value (measured by the last cell in the first row)    // creates a 'colnum' variable and sets it to the number of the last column in the first row
		
			int y=0;										//creates a 'y' variable and sets it to 0 
			Row row = ws.getRow(y);							//creates a Row variable called 'row' and sets it to the row that is stated in 'y'
				for(int x=0; x<colNum ; x++){				//starts a for loop, creats 'x' variable, for as long as x is less than 'colNum' run this block of code, add 1 to x everytime the code is run
					Cell cell = row.getCell(x);				//create a cell variable called 'cell' set it to the cell 'x'(x will grow each time the loop runs) in 'row'
					
					String value = cellToString(cell);			//finds the string that is in each cell by calling the cellToString method, cellToSting(a cell variable that you wish to convert to a string);
					if (Objects.equals(value, username)){		// if the string in the cell is the user name run the following code
						result= x;								//sets the 'result' variable to whatever 'x' is
						break;									//breaks out of the loop, we have found the x coordinate so there is no reason to continue 
					}											//if the string in the cell does not match the username continue looping
				}
				if (result==0){
					System.out.println("user not found");
				}
				return result;		//return the 'result' variable to whoever called this method
		}
		
	public static int findY(int x) throws IOException, EncryptedDocumentException, InvalidFormatException{		//find the y coordinate by searching the first column for the current date
		String currentDate= getDate();							//calls the getDate() method and save it to 'currentDate'
		int result = 0;											//creates the 'result' variable and sets it to 0 
		File fileName = new File ("Activity Log.xls");			//sets the file name
		FileInputStream fis = new FileInputStream(fileName);	//create a file input stream to read the excel file
		HSSFWorkbook wb = new HSSFWorkbook(fis);				//create a work book from the file input stream
		HSSFSheet ws = wb.getSheet("sheet1");					//get the sheet called 'sheet1' from the file input stream

		int rowNum = ws.getLastRowNum() ;			//rowNum is the max y value. POI is 0 based		//get the highest y value with text in the cell (the last date an entry was made)
	
	
		for(int y=0; y<=rowNum; y++){							//start a for loop, create a 'y' variable, for as long as 'y' is less than 'rowNum' run this block of code, and 1 to 'y' everytime this code is run 
			HSSFRow row = ws.getRow(y);							//create a Row variable called 'row' from the worksheet with the 'y' value ('y' grows everytime)
			HSSFCell cell = row.getCell(0);						//looks at the first cell in the 'row' variable
			String value = cellToString(cell);					//calls the cell to string method, cellToSting(a cell variable that you wish to convert to a string);
			if (Objects.equals(value,currentDate )){			//if the value of the cell equals the current date run the following code
				 result = y;									//set 'result' to what ever 'y' is
				break;											//break out of the loop, we found the row that we need 
			}
			if(y==rowNum){														//if we loop through all of the rows and have not found the current date run this code	
				FileInputStream inp = new FileInputStream(fileName);			//create a new input stream to read from
			    Workbook wbf = WorkbookFactory.create(inp);						//create a workbook with the workbook factory method
			    FileOutputStream fileOut = new FileOutputStream(fileName);		//create a file output stream to write to the excel file
			    Sheet sheet = wbf.getSheet("sheet1");							//get the sheet called sheet one from the 'wbf' workbook
			    writeCell(sheet,0,rowNum+1, getDate());							//calls the writeCell method, the y value is 'rowNum' +1 to write in a new row
			     
			    
			    wbf.write(fileOut);		//write to the excel file
			    fileOut.close();		//close the output file stream
				
			    result = rowNum+1;		//set the result to 'rowNum' + 1
				break;					//break out of the loop
				
			}
		}
		wb.close();				//close the workbook at the start of the for loop
		return result;			//return the result to whoever called this method
	}
	@SuppressWarnings("deprecation")
	public static String cellToString(Cell cell){		//converts a cell to a string, cellToString(the cell that you are referencing);
		int type;										//creates a integer called 'type'
		Object result;									//creates a result object
		type=cell.getCellType();						//set 'type' to the cell type
		
		switch (type){			//switch between the possible outcomes of 'type'
		
		case 0 : 								//numeric value		//if the type is a number set result to the numeric cell value
			result= cell.getNumericCellValue();
			break;								//break the switch loop
		case 1:									//String value			//if the type is a string set result to the string cell value
			result= cell.getStringCellValue();
			break;								//break the switch loop
		default:								//the default is if 'type' fits none of the categories above
			throw new RuntimeException("This type of cell is not supported!");		//throws an exception
		}
		return result.toString();			//return the object variable 'result' as a string value
	}
	public static void writeCell(Sheet sheet, int x, int y, String cellContents){		//writes contents to a cell, writeCell(name of the sheet, x coordinate , y coordinate , text to be written);
		System.out.println("cell contents:"+cellContents + " x y cords:"+x+" , "+y+" sheet: "+sheet);
		Row r = sheet.getRow(y);			//create a Row variable called 'r' and set it to the row 'y' in the Sheet 'sheet' 
		if (r==null){						//if the row given is blank this creates the row to be able to write to it
			r= sheet.createRow(y);
		}
		
		Cell c = r.getCell(x);					//creates a Cell variable called 'c' and gets the cell that is in row 'r' and is cell number 'x'
		if(c == null){							//if the cell is empty this creates it then writes to it
			c= r.createCell(x);
			c.setCellValue(cellContents);		//sets the value of the cell to whatever was passed to the method in 'cellContents'
		}
		else if(c != null){											//if the cell is not empty this saves what is in it
			String currentContents = c.getStringCellValue();		//saves what is currently in the cell in 'currentContents'
			c.setCellValue(currentContents +" "+ cellContents);		//writes to the cell, 'currentContents' + 'cellContents'		
		}
	}
	public static void setCell(Sheet sheet, int x, int y, String cellContents){		//writes contents to a cell, writeCell(name of the sheet, x coordinate , y coordinate , text to be written);
		Row r = sheet.getRow(y);		//create a Row variable called 'r' and set it to the row 'y' in the Sheet 'sheet' 
		if (r==null){					//if the row given is blank this creates the row to be able to write to it
			r= sheet.createRow(y);
		}
		
		Cell c = r.getCell(x);					//creates a Cell variable called 'c' and gets the cell that is in row 'r' and is cell number 'x'
		if(c == null){							//if the cell is empty this creates it then writes to it
			c= r.createCell(x);
			c.setCellValue(cellContents);		//sets the value of the cell to whatever was passed to the method in 'cellContents'
		}
		else if(c != null){
			c.setCellValue(cellContents);			
		}
	}	
	public static void playSound(URL Sound){		//plays a sound from the system, playSound (URL of the sound you want to play)
		
		try {
			Clip clip = AudioSystem.getClip();
			clip.open(AudioSystem.getAudioInputStream(Sound));
			clip.start();
		} catch (LineUnavailableException | IOException | UnsupportedAudioFileException e) {
			e.printStackTrace();
		}
		
	}
	
	public static float hourComp(String login, String logout){  		//calculates the numbers of hours between logged times (time of login ex: "11,25" , time of logout ex: "15,39")
	int loginHrs=Integer.parseInt(login.substring(0,2));				//gets the hour number from the login time (ex: 11)
	int loginMin=Integer.parseInt(login.substring(3,5));				//gets the minute number from the login time (ex: 25)
	int logoutHrs=Integer.parseInt(logout.substring(1,3));				//gets the hour number for the logout time (ex: 15)	
	int logoutMin=Integer.parseInt(logout.substring(4,6));				//gets the minute number for the logout time (ex: 39)
	int totalLoginMin=(loginHrs*60)+loginMin;							//finds the total minute by multiplying the hours by 60 and adding the min
	int totalLogoutMin=(logoutHrs*60)+logoutMin;						//same as that ^ but for logout time
	float diff = (totalLogoutMin-totalLoginMin);						//finds the difference by subtracting
	float result = diff/60;												//Divides by 60 to find hours, and returns the value to who ever called this method 
	return result;
	}
	
	public static void shutdownCheck() throws EncryptedDocumentException, InvalidFormatException, IOException{			//Shutdown checks, logs anyone out who is logged in when you close the program, 
		Set set = objMap.entrySet();																					//this prevents someone having a login time with out a log out time
	
											//get an iterator
		Iterator i = set.iterator();		//set an iterator
		int userLoopID;						//create a 'userLoopID' variable and set it to 0
		String userLoopName;
		
		while(i.hasNext()){									//start the while loop, while there are still more people in the hashmap loop this code
			Map.Entry me = (Map.Entry)i.next();				//create a map entry of 'i'
			Students userLoop = objMap.get(me.getKey());	//create a JLable variable called 'userLoop' and set it to the key of the current map entry
			if (userLoop.state){							//if that user is enabled run this code
				userLoopName = userLoop.name;				//create a 'userLoopName' variable and get the name from the 'userLoop'
				userLoopID = userLoop.studentID;	
				
				userLoop.state = false;							//set the JLabel 'userLoop' to disabled (logged out)
				writeFile(userLoopName , false ,userLoopID );	//write the change to the excle file
						}
					}
		System.exit(0);			//close the program  
			}
}
