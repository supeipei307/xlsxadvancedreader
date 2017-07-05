package com.pairdata.xlsx;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.SAXException;

public class XlsxReader {
	private OPCPackage opcPackage;
	private ReadOnlySharedStringsTable sharedStringsTable;
	private Iterator<InputStream> sheetInputSteams;
	private InputStream sheetInputSteam;
	private XMLStreamReader sheetXmlReader;
	private ArrayList<Integer> numFmtIdArray;

	private int lineNumber = 0;
	private int cellCount = 0;
	
    public XlsxReader(String excelFilePath, int sheetNum, int cellCount){
    	try {
    		this.cellCount = cellCount;
    		
			opcPackage = OPCPackage.open(excelFilePath, PackageAccess.READ);
			sharedStringsTable = new ReadOnlySharedStringsTable(opcPackage);
			
			XSSFReader xssfReader = new XSSFReader(opcPackage);
			XMLInputFactory factory = XMLInputFactory.newInstance();
			
			InputStream stylesInputSteam = xssfReader.getStylesData();
			XMLStreamReader stylesXmlReader = factory.createXMLStreamReader(stylesInputSteam);
		    initStyles(stylesXmlReader);
		    
		    sheetInputSteams = xssfReader.getSheetsData();
		    for (int i=0; i<sheetNum; i++)
		    	sheetInputSteam = sheetInputSteams.next();
		    
		    if (sheetInputSteam != null){
		        sheetXmlReader = factory.createXMLStreamReader(sheetInputSteam);
		        while (sheetXmlReader.hasNext()) {
		        	sheetXmlReader.next();
		        	if (sheetXmlReader.isStartElement() && sheetXmlReader.getLocalName().equals("sheetData"))
		        	    break;
		        }
		    }
		} catch (InvalidOperationException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (OpenXML4JException e) {
			e.printStackTrace();
		} catch (XMLStreamException e) {
			e.printStackTrace();
		} catch (SAXException e) {
			e.printStackTrace();
		}
    }
    
    public void setCellCount(int cellCount){
    	this.cellCount = cellCount;
    }
    
    public int getLineNumber(){
    	return lineNumber;
    }
    
    public List<String> readLine() throws XMLStreamException{
    //public String[] readLine() throws XMLStreamException{
    	boolean isInRow = false;
    	List<String> lineArrayList = new ArrayList<String>();
    	
    	while (sheetXmlReader.hasNext()){
    		sheetXmlReader.next();
    		
    		if (sheetXmlReader.isStartElement() && sheetXmlReader.getLocalName().equals("row"))
    			isInRow = true;
    		else if (sheetXmlReader.isEndElement() && sheetXmlReader.getLocalName().equals("row")){
    			while (lineArrayList.size() < cellCount)
    				lineArrayList.add("");
    			lineNumber++;
    			return lineArrayList;
    			//return lineArrayList.toArray(new String[lineArrayList.size()]);
    		}
    		else if (isInRow && sheetXmlReader.isStartElement() && sheetXmlReader.getLocalName().equals("c")){
    			CellReference cellReference = new CellReference(sheetXmlReader.getAttributeValue(null, "r"));
    			while (lineArrayList.size() < cellReference.getCol())
    			    lineArrayList.add("");
    			String cellType = sheetXmlReader.getAttributeValue(null, "t");
    			String cellStyle = sheetXmlReader.getAttributeValue(null, "s");
    			//System.out.print(cellType + " - " + cellStyle + " ");
    			
    			lineArrayList.add(getCellValue(cellType, cellStyle));
    		}
    	}
    	return null;
    	//return lineArrayList.toArray(new String[lineArrayList.size()]);
    }

	private String getCellValue(String cellType, String cellStyle) throws XMLStreamException {
		while (sheetXmlReader.hasNext()) {
		    sheetXmlReader.next();
		    if (sheetXmlReader.isStartElement() && sheetXmlReader.getLocalName().equals("v")){
		    	if (cellType != null && cellType.equals("s")){
		    	    int idx = Integer.parseInt(sheetXmlReader.getElementText());
		    	    return new XSSFRichTextString(sharedStringsTable.getEntryAt(idx)).toString();
		    	}
		    	else if (cellType != null && cellType.equals("b")){
		    		return (Integer.parseInt(sheetXmlReader.getElementText()) == 1)?"True":"False";
		    	}
		    	else if (cellStyle != null){
		    		int idx = Integer.parseInt(cellStyle);
		    		int numFmtId = numFmtIdArray.get(idx);
		    		
		    		if (numFmtId >= 14 && numFmtId <= 17){
		    			Date date = DateUtil.getJavaDate(Double.parseDouble(sheetXmlReader.getElementText()));
		    			SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
		    			return simpleDateFormat.format(date);
		    		}
		    		else if ((numFmtId >= 18 && numFmtId <= 21) || (numFmtId >= 45 && numFmtId <= 47)){
		    			Date date = DateUtil.getJavaDate(Double.parseDouble(sheetXmlReader.getElementText()));
		    			SimpleDateFormat simpleDateFormat = new SimpleDateFormat("HH:mm:SS");
		    			return simpleDateFormat.format(date);
		    		}
		    		else if (numFmtId == 22){
		    			Date date = DateUtil.getJavaDate(Double.parseDouble(sheetXmlReader.getElementText()));
		    			SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:SS");
		    			return simpleDateFormat.format(date);
		    		}
		    	}
		    	return sheetXmlReader.getElementText();
		    }
		}
		return null;
	}

	private void initStyles(XMLStreamReader stylesXmlReader) throws NumberFormatException, XMLStreamException {
		boolean isInCellXfs = false;
		this.numFmtIdArray = new ArrayList<Integer>();
		
		if(stylesXmlReader != null){
			while(stylesXmlReader.hasNext()) {
				stylesXmlReader.next();
				
				if(stylesXmlReader.isStartElement() && stylesXmlReader.getLocalName().equals("cellXfs"))
					isInCellXfs = true;
				else if(stylesXmlReader.isEndElement() && stylesXmlReader.getLocalName().equals("cellXfs")){
					isInCellXfs = false;
					break;
				}
				else if(isInCellXfs && stylesXmlReader.isStartElement() && stylesXmlReader.getLocalName().equals("xf"))
				    numFmtIdArray.add(Integer.parseInt(stylesXmlReader.getAttributeValue(null, "numFmtId")));
		    }
		}
	}
	
	@Override
	protected void finalize() throws Throwable {
	    if (opcPackage != null)
		    opcPackage.close();
	    super.finalize();
	}
}