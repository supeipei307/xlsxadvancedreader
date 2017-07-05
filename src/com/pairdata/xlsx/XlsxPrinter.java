package com.pairdata.xlsx;

import java.util.ArrayList;
import java.util.List;

import javax.xml.stream.XMLStreamException;

public class XlsxPrinter {
	
	//args[0]: Excel file path
	//args[1]: Sheet number, start from 1
	//args[2]: Max cell count, input 0 means determine by the first line
	//args[3]: Delimiter in the output (ASCII code, for example, TAB is 9)
	//args[4]: Which line to start to read, 0 or 1 means start from the first line
	//args[5]: Which line to end the reading, 0 means read to the end of the sheet
	public static void main(String[] args) {
		
		int sheetNum = Integer.parseInt(args[1]);
		int maxCellCount = Integer.parseInt(args[2]);
		char delimiter = (char)Integer.parseInt(args[3]);
		int startLineNumber = Integer.parseInt(args[4]);
		int endLineNumber = Integer.parseInt(args[5]);
		
		XlsxReader xlsxReader = new XlsxReader(args[0], sheetNum, maxCellCount);
		List<String> list = new ArrayList<String>();
		
		try {
			do{ 
				list = xlsxReader.readLine();
				if (list == null)
					break;
				
				int lineNumber = xlsxReader.getLineNumber();
				
				if(endLineNumber > 0 && endLineNumber <= lineNumber)
					break;
				
				if (maxCellCount <= 0 && lineNumber == 1)
					xlsxReader.setCellCount(list.size());
				
				if (startLineNumber > lineNumber)
					continue;
				
				for (int i=0; i<list.size(); i++){
					System.out.print(list.get(i));
					if (i<list.size() - 1)
						System.out.print(delimiter);
				}
				System.out.println("");
			}while(list != null);
		} catch (XMLStreamException e) {
			e.printStackTrace();
		}
	}

}
