package HybridFWK;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class TestDataLoader {
		
	static XSSFSheet TestCases,orgDetails,vmdDetails;
	static XSSFWorkbook wb ;
	ArrayList<ArrayList<String>> excel_data = new ArrayList<ArrayList<String>>();
	ArrayList<String> sheetnames;
	HashMap<String, List<String>> Testconfig = new HashMap<String, List<String>>();
	int suiteCount = 0 ,dataCount=0;
	List<String> datalist=new ArrayList<String>();
	List<String> suitelist=new ArrayList<String>();
	String value,key;
	
	public  void DataLoader() throws IOException {
	
	FileInputStream fis = new FileInputStream("E:\\OrgDeletion\\TestData.xlsx");
	wb = new XSSFWorkbook(fis);
	int sheetcount = wb.getNumberOfSheets();
	sheetnames = new ArrayList<String>();	
	System.out.println("sheetcount: "+sheetcount);
	
    for(int i=0; i<sheetcount;i++)
    {
       String currentSheetname=wb.getSheetName(i);
       if(currentSheetname.toLowerCase().contains("config"))
	    {
    	    System.out.print("i have Test Config" +"\n");
    	    
    	    for (int counter = 0; counter < getSheetdata(i).size(); counter++)
    	    {   	      	    
    	      for (int k = 0; k<getSheetdata(i).get(0).size(); k++)
    	      {
    	    	  String text = getSheetdata(i).get(counter).get(k);
    	    	  
    	          switch (text.toLowerCase()) {
	    	          case "suite":	    	        	    
	    	        	    suiteCount++;	    	        	    
	    	        	    break;
	    	          case "data":
	    	        	    dataCount++;	    	        	  
	    	        	    break;    
               }    	          
    	     }
    	    }	 
    	    	 
	    int rows =wb.getSheetAt(i).getPhysicalNumberOfRows();
	    for(int j=1; j<=rows; j++)
	    	 {    			
	  			  Row row = wb.getSheetAt(i).getRow(j);
	  			  
	  			  if(row !=null)
	  			  {
	  			  Cell valueCell =row.getCell(0); 
	  			  Cell keyCell = row.getCell(1);
	  			  String value = valueCell.getStringCellValue().trim();
	  			  String key = keyCell.getStringCellValue().trim(); 
	  			  System.out.println("Key+Value "  + key+" + " + value );
	  			/**
	  	            if(Testconfig.get((key))==null)
	  	            {
	  	            	Testconfig.put(key,new  ArrayList<String>(Arrays.asList(value)));
	  	               
	  	            }
	  	            else
	  	            {
	  	            	Testconfig.get(key).addAll(new  ArrayList<String>(Arrays.asList(value)));
	  	               
	  	            }
	  	        
	  			  }	**/
	  			  
	   	       
	  			   			  
	  	      if(suiteCount!= 0 && key.equals("SUITE"))
	  			{
	  				suitelist.add(value);
	  				suiteCount--;
	  				Testconfig.put(key.toLowerCase(), suitelist); 
	  			 }
	  	      if(dataCount!=0 && key.equals("DATA"))
	  			 {
	  			    datalist.add(value);
	  				dataCount--;
	  				Testconfig.put(key.toLowerCase(), datalist);
	  			 }  		 			 
	  			} 			 	
	  		  }	     
	          sheetnames.add(wb.getSheetName(i));   
	    }	   
	  }  
	}
	  
    
	public  ArrayList<ArrayList<String>> getSheetdata(int i) {
    	
    excel_data = new ArrayList<ArrayList<String>>();
    XSSFSheet sheet = wb.getSheetAt(i);   
    DataFormatter dataFormatter = new DataFormatter();    
    Iterator<Row> rowIterator = sheet.rowIterator();
    while (rowIterator.hasNext()) {
    	ArrayList<String> filedata = new ArrayList<String>() ;
        Row row = rowIterator.next();        
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String cellValue = dataFormatter.formatCellValue(cell);
            //System.out.print(cellValue + "\t");
            filedata.add(cellValue);           
            
        }
        excel_data.add(filedata);       
        
    }
    return excel_data;
    
    }
    
    public void dataPrint(){
    	
    	//System.out.println("Excel Data"  +excel_data );
    	  System.out.println("TestConfig" + Testconfig);
    	
    }
    
    public static void main(String[] args) throws IOException {
    	TestDataLoader tdl = new TestDataLoader();
    	tdl.DataLoader();
    	tdl.dataPrint();
    	
    }
}
	
