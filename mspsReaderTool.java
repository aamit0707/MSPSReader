import java.awt.Button;
import java.awt.Color;
import java.awt.FileDialog;
import java.awt.Font;
import java.awt.Frame;
import java.awt.Label;
import java.awt.TextField;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DecimalFormat;
import java.util.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class mspsReaderTool extends KeyAdapter implements ActionListener, WindowListener
{

   	 Connection conn;
         Statement  st;
	 ResultSet rs;

	String targetFilePath, sourceFilePath;
	Frame f;
	TextField tfilepath, targetfilepath;
	Button c2,bsignout,bcompute, acnpath, targetpathButton;
	Label l,path,targetPath,message,tex,show,review,ctpr,ctex,rework,coreview,corework,cdreview,cdrework;
	Label lpatherror,tpatherror,ltex,lshow,lreview,lctcs,lctex,lrework;
	Label lev,lpv,lac,lbac,lev1,lpv1,lac1,lbac1,letc,letc1,lspe,lspe1,lcoreview,lcorework,lcdreview,lcdrework; 
	Label ttpm,ltpr,ttex,treview,tctpr,tctex,trework;

 	  
				mspsReaderTool()
				{	
					f=new Frame("Main Window");
					Color c1=new Color(155,208,251); 
					Color c2=new Color(170,208,255);
					
					f.setBackground(c1);       
					f.setSize(900,400);
					Font f2 = new Font("Monotype Corsiva", Font.BOLD, 20);
					Font f4 = new Font("Monotype Corsiva", Font.PLAIN, 13);
					Font j = new Font("Monotype Corsiva", Font.BOLD, 20);
					Font j1 = new Font("Monotype Corsiva", Font.BOLD, 40);
					
					l = new Label("----- MSPS Reader Tool ----- ");
					l.setFont(j1);
					l.setForeground(Color.black);
					f.add(l);
					l.setBounds(200, 40, 800, 50);
					
					path = new Label("ACN Path");
					path.setFont(j);
					path.setForeground(Color.black);
					f.add(path);
					path.setBounds(25, 120, 120, 25);
					
					acnpath = new Button("Open");
					acnpath.setFont(f2);
					acnpath.setBackground(c2);
					f.add(acnpath);
					acnpath.setBounds(770, 120, 100, 30);
					
					targetPath = new Label("Target File");
					targetPath.setFont(j);
					targetPath.setForeground(Color.black);
					f.add(targetPath);
					targetPath.setBounds(25, 200, 120, 25);

					targetpathButton = new Button("Open");
					targetpathButton.setFont(f2);
					targetpathButton.setBackground(c2);
					f.add(targetpathButton);
					targetpathButton.setBounds(770, 200, 100, 30);
									 
			        
					lpatherror = new Label("");
					lpatherror.setFont(j);
					lpatherror.setForeground(Color.red);
					f.add(lpatherror);
					lpatherror.setBounds(150, 160, 550, 25);
					
					tpatherror = new Label("");
					tpatherror.setFont(j);
					tpatherror.setForeground(Color.red);
					f.add(tpatherror);
					tpatherror.setBounds(150, 240, 550, 25);
					
					message = new Label("");
					message.setFont(j);
					message.setForeground(Color.red);
					f.add(message);
					message.setBounds(150, 340, 550, 25);
					
					
					bcompute = new Button("Compute");
					bcompute.setFont(f2);
					bcompute.setBackground(c2);
					f.add(bcompute);
					bcompute.setBounds(770, 260, 100, 30);
					bcompute.addActionListener(this); 
					
					
					tfilepath = new TextField("");
					tfilepath.setFont(f4);
					tfilepath.setBackground(c2);
					tfilepath.setBounds(150, 120, 600, 25);
					f.add(tfilepath); 
					
					targetfilepath = new TextField("");
					targetfilepath.setFont(f4);
					targetfilepath.setBackground(c2);
					targetfilepath.setBounds(150, 200, 600, 25);
					f.add(targetfilepath);
					
					
					bsignout = new Button("Close");
					bsignout.setFont(f2);
					bsignout.setBackground(c2);
					//f.add(bsignout);
			        bsignout.setBounds(800, 50, 90, 40);  
			    
			        bsignout.addActionListener(this);    
			        acnpath.addActionListener(this); 
			        targetpathButton.addActionListener(this);
			        f.addWindowListener(this);
					           
					f.setLayout(null);
					f.setResizable(false);
				    f.setVisible(true);
					
									
					
				}
				
				public void actionPerformed(ActionEvent e)
		        {
			                if(e.getSource() ==bsignout )
			                  {
			                       System.exit(1);
			                  }
			                else if(e.getSource()==acnpath)
			                  {
			                	lpatherror.setText("");	
			                	tfilepath.setText("");
			                	message.setText("");
			                	FileDialog fd=new FileDialog(f,"OPEN",FileDialog.LOAD);
			            		fd.setVisible(true);
			            		String fi,di;
			            		fi=fd.getFile();
			            		di=fd.getDirectory();
			            		String path,temp;
			            		temp=di.concat("\\");
			            		sourceFilePath=temp.concat(fi);
			            		tfilepath.setText(sourceFilePath);
			                  }
			                
			                
			                else if(e.getSource()==targetpathButton)
			                  {
			                	lpatherror.setText("");	
			                	targetfilepath.setText("");
			                	message.setText("");
			                	FileDialog fd=new FileDialog(f,"OPEN",FileDialog.LOAD);
			            		fd.setVisible(true);
			            		String fi,di;
			            		fi=fd.getFile();
			            		di=fd.getDirectory();
			            		String path,temp;
			            		temp=di.concat("\\");
			            		targetFilePath=temp.concat(fi);
			            		targetfilepath.setText(targetFilePath);
			                  }
			                
			                 else if(e.getSource()==bcompute)
			                  {
			                	 	message.setText("Process is in Progess...");
			                	 	String sourcefile = tfilepath.getText();
			                	 	String targetfile = targetfilepath.getText();
				                	if(!sourcefile.contains(".xlsx"))
				                	{
				                		lpatherror.setText("File format is not .xlsx");
				                	}
				                	
				                	if(!targetfile.contains(".xlsx"))
				                	{
				                		//System.out.println("test "+targetFilePath+" test "+targetfile);
				                		tpatherror.setText("File format is not .xlsx");
				                	}
				                	
				                	else
				                	{
				                		//File excel = new File("D:\\Studies\\MSPS_Reader\\ACN Task Resource Usage Per Period (Task Level) Report (5).xlsx");
				                		File excel = new File(sourceFilePath);
				                		FileInputStream fis = null;
										try 
										{
											fis = new FileInputStream(excel);
										} 
										catch (FileNotFoundException e1) 
										{
											// TODO Auto-generated catch block
											message.setText("File is already Open");
											e1.printStackTrace();
										}
				                		
										XSSFWorkbook wb = null;
										try 
										{
											wb = new XSSFWorkbook(fis);
										} 
										catch (IOException e1) 
										{
											// TODO Auto-generated catch block
											e1.printStackTrace();
										}
				                		
				                		XSSFSheet sheet = wb.getSheet("Sheet2");
				                		
				                		int rowNum = sheet.getLastRowNum()+1;
				                		System.out.println("Total Rows - "+rowNum);
				                		int colNum = sheet.getRow(1).getLastCellNum();
				                		System.out.println("Total Column - "+colNum);
				                		String[] projectdata = new String[rowNum];
				                		String[] resourcedata = new String[rowNum];
				                		int projectCount=0;
				                		for(int i=1;i<rowNum-1;i++)
				                		{
				                				int j=0;
				                				XSSFRow row = sheet.getRow(i);
				                				XSSFCell cell=row.getCell(j);
				                				String value = cell.getStringCellValue();
				                				boolean b;
				                				b=value.startsWith("SVT");
				                				if(b==true)
				                					{
				                						
				                						projectdata[projectCount]=  value;
				                						//System.out.println(data[projectCount][j]);
				                						projectCount++;
				                					}
				                		} //project name
				                		int resourceCount=0;
				                		for(int i=3;i<rowNum-1;i++)
				                		{
				                			int j=1;
				                			XSSFRow row = sheet.getRow(i);
				                			XSSFCell cell=row.getCell(j);
				                			String value = cell.getStringCellValue();
				                			//boolean b;
				                			//b=value.startsWith(" ");
				                	    	//if(b!=true)
				                				//{
				                					
				                					resourcedata[resourceCount]=  value;
				                					//System.out.println(resourcedata[resourceCount][0]);
				                					resourceCount++;
				                				//}
				                		} //resource name
				                		
				                		String[] resourceDistinctData = new HashSet<String>(Arrays.asList(resourcedata)).toArray(new String[0]);
				                		String[] projectDistinctData = new HashSet<String>(Arrays.asList(projectdata)).toArray(new String[0]);
				                		System.out.println("Unique Resource Count - "+resourceDistinctData.length);
				                		System.out.println("Unique Project Count - "+projectDistinctData.length);
				                		
				                		
				                		String[][] finalDataSheet = new String[rowNum][3];
				                	//	String[][] superDataSheet = new String[rowNum][3];
				                		int tempResource = 0;
				                		double totalHrsByProject = 0.0;
				                		String prevProjectName = "";
				                		String prevResourceName = "";
				                		/*for(int i=0;i<resourceDistinctData.length;i++)
				                		{
				                				for(int j=2;j<rowNum-1;j++)
				                				{
				                					
				                					int p=2;
				                					int h=9;
				                					int r=1;
				                					XSSFRow row = sheet.getRow(j);
				                					XSSFCell cell=row.getCell(p);
				                					XSSFCell cellResourceName=row.getCell(r);
				                					//XSSFCell cellResourceNameCheck=row.getCell(r-1);
				                					XSSFCell cellActualHours=row.getCell(h);
				                					String currProjectName = cell.getStringCellValue();
				                					String resourceName = cellResourceName.getStringCellValue();
				                					double actualTemp=(double) cellActualHours.getNumericCellValue();
				                					if (resourceName == resourceDistinctData[i])
				                					{
				                						if(actualTemp!=0.0)
				                						{
				                							if(prevProjectName == "" || currProjectName == prevProjectName)
				                							{	
				                							
				                								prevProjectName=currProjectName;
				                								prevResourceName = resourceName;
				                								totalHrsByProject = totalHrsByProject + actualTemp;
				                							}
				                							else{
				                								finalDataSheet[tempResource][0]= prevResourceName;
				                								finalDataSheet[tempResource][1]= prevProjectName;
				                								finalDataSheet[tempResource][2]= Double.toString(totalHrsByProject);
				                								prevProjectName = "";
				                								prevResourceName = "";
				                								totalHrsByProject = 0.0;
				                								tempResource++;
				                								 j = j-1;
				                							}				                						
				                						}

				                					}
				                					
				                				}
				                				
				                		}		// resource wise total sum
*/				                		
				                		
				                		String[][] finalDataSheetCheck = new String[rowNum][3];
					                	//	String[][] superDataSheet = new String[rowNum][3];
					                	int tempResourceCheck = 0;
				                		double totalHrsByProjectCheck = 0.0;
				                		String prevProjectNameCheck = "";
				                		String prevResourceNameCheck = "";
				                		
				                		for(int j=2;j<rowNum-1;j++)
		                				{
		                					
		                					int p=2;
		                					int h=9;
		                					int r=1;
		                					XSSFRow row = sheet.getRow(j);
		                					XSSFCell cell=row.getCell(p);
		                					XSSFCell cellResourceName=row.getCell(r);
		                					XSSFCell cellActualHours=row.getCell(h);
		                					String currProjectNameCheck = cell.getStringCellValue();
		                					String currresourceNameCheck = cellResourceName.getStringCellValue();
		                					double actualTemp=(double) cellActualHours.getNumericCellValue();
		                					
		                					
		                					if(actualTemp!=0.0)
	                						{
			                					if(prevResourceNameCheck == "" || prevResourceNameCheck == currresourceNameCheck)
			                							{	
					                						prevProjectNameCheck=currProjectNameCheck;
			                								prevResourceNameCheck = currresourceNameCheck;
			                								totalHrsByProjectCheck = totalHrsByProjectCheck + actualTemp;
			                								//System.out.println("Check = "+j);
			                							}
	                							else
		                								{
			                								finalDataSheetCheck[tempResourceCheck][0]= prevResourceNameCheck;
			                								finalDataSheetCheck[tempResourceCheck][1]= prevProjectNameCheck;
			                								finalDataSheetCheck[tempResourceCheck][2]= Double.toString(totalHrsByProjectCheck);
			                								prevProjectNameCheck = "";
			                								prevResourceNameCheck = "";
			                								totalHrsByProjectCheck = 0.0;
			                								tempResourceCheck++;
			                								System.out.println("Check Here= "+j);
			                								j = j-1;
			                								
			                							}				                						
	                						 }	                				 
		                				}
				                		
				                		int finalSheetRowCount=0;
				                		for (int i = 0; i < finalDataSheetCheck.length; i++)
				                	    {
				                			if(finalDataSheetCheck[i][0]!=null)
				                			{
				                				finalSheetRowCount++;
				                			}
				                	    }
				                		System.out.println("Number of Lines "+finalSheetRowCount);
				                		
			                		
				                		String[][] superDataSheet = new String[finalSheetRowCount+1][3];
				                		superDataSheet[0][0]="Resource Name";
	                	            	superDataSheet[0][1]="Project Name";
	                	            	superDataSheet[0][2]="Actual Charged Hours";
				                		
				                		for (int i = 1; i <= finalSheetRowCount; i++)
				                	    {
				                			superDataSheet[i][0]=finalDataSheetCheck[i-1][0];
		                	            	superDataSheet[i][1]=finalDataSheetCheck[i-1][1];
		                	            	superDataSheet[i][2]=finalDataSheetCheck[i-1][2];
				                	    }
				                		System.out.println(superDataSheet.length);
				                		/*double sum=0;
				                		int countUnique =0;
				                		for (int j = 0; j < 5; j++)
				                	    {
				                	        for (int k = j + 1; k < 15; k++) 
				                	        {
				                	            if (k != j && finalDataSheet[k][0] == finalDataSheet[j][0] && finalDataSheet[k][1] == finalDataSheet[j][1])
				                	            {
				                	                sum = sum + Double.parseDouble(finalDataSheet[k][2]);
				                	                //System.out.println("Duplicate found: " + array[k] + " " + "Sum of the duplicate value is " + sum);
				                	            }
				                	            if (k != j && finalDataSheet[j][0] != finalDataSheet[k][0] && finalDataSheet[j][1] != finalDataSheet[k][1])
				                	            {
				                	            	superDataSheet[countUnique][0]=finalDataSheet[j][0];
				                	            	superDataSheet[countUnique][1]=finalDataSheet[j][1];
				                	            	superDataSheet[countUnique][2]=Double.toString(sum);
				                	            	sum=0;
				                	            	countUnique++;
				                	            	
				                	            }
				                	        }
				                	    }
				                		
				                		
				                		for(int i=0;i<2;i++)
				                		{
				                			boolean b= superDataSheet[i][1].startsWith("SVT");
				                			if(b==true)
				                			{
				                				System.out.println(superDataSheet[i][0]+"  "+superDataSheet[i][1]+"   "+superDataSheet[i][2]);
				                			}
				                		}
				                		
				                		*/
				                						                		
				                		XSSFWorkbook workbook = new XSSFWorkbook();
				                        XSSFSheet sheetWrite = workbook.createSheet("Java Books");
				                        
				                        Arrays.sort(superDataSheet, new Comparator<String[]>()
				                        {
				                    	    @Override
				                    	    public int compare(String[] first, String[] second){
				                    	        // compare the first element
				                    	        int comparedTo = first[0].compareTo(second[0]);
				                    	        // if the first element is same (result is 0), compare the second element
				                    	        if (comparedTo == 0) return first[1].compareTo(second[1]);
				                    	        else return comparedTo;
				                    	    }
				                    	}
				                        );
				                		
				                        int rowCount = 0;
				                        
				                        for (Object[] aBook : superDataSheet) 
				                        {
				                            Row row = sheetWrite.createRow(rowCount++);
				                             
				                            int columnCount = 0;
				                             
				                            for (Object field : aBook) {
				                                Cell cell = row.createCell(columnCount++);
				                                if (field instanceof String) {
				                                    cell.setCellValue((String) field);
				                                } else if (field instanceof Integer) {
				                                    cell.setCellValue((Integer) field);
				                                }
				                            }
				                             
				                        }
				                         
				                         //targetFilePath  "C:\\Amit\\JavaBooks3.xlsx"
				                        try (FileOutputStream outputStream = new FileOutputStream(targetFilePath,false)) 
				                        {
				                            workbook.write(outputStream);
				                            
				                            message.setText("Target File is Ready");
				                            
				                            outputStream.close();
				                        } 
				                        catch (IOException e1) 
				                        {
				                        	message.setText("Please close the Target File First");
											// TODO Auto-generated catch block
											e1.printStackTrace();
										}
				                	}
			                  }
				                	
				                	
				                	}
				                		

				  		                               
				public void windowClosing(WindowEvent we)
				{
			  		System.out.println("Window close");
					we.getWindow().dispose();
			
				}
			  	public void windowActivated(WindowEvent we)
				{
				}
				public void windowClosed(WindowEvent we)
				{
					we.getWindow().dispose();
				}
				public void windowDeactivated(WindowEvent we)
				{
					System.out.println("Window Deactivated");
				}
				public void windowDeiconified(WindowEvent we)
				{
				}
				public void windowIconified(WindowEvent we)
				{
				}
				public void windowOpened(WindowEvent we)
				{
					System.out.println("Window Opened");
				}
				
					
				public static void main(String s[]) throws IOException
				{
					mspsReaderTool p=new mspsReaderTool();  
				}

}