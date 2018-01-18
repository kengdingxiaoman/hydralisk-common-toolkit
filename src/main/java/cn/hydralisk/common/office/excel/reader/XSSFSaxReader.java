package cn.hydralisk.common.office.excel.reader;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 *
 * @author master.yang
 */
public class XSSFSaxReader implements IExcelReader {

	private InputStream inputStream;
	private OPCPackage opcPackage;
	private int startRow;
	private int endRowIgnoredNum;
	
	public XSSFSaxReader(InputStream inputStream, int startRow, int endRowIgnoredNum) {
		this.inputStream = inputStream;
		this.startRow = startRow;
		this.endRowIgnoredNum = endRowIgnoredNum;
	}

	@Override
	public void read(RowReader reader) throws ExcelReadException {
		try {
			this.opcPackage = OPCPackage.open(inputStream);
			XSSFReader r = new XSSFReader(opcPackage);
			
			XMLReader parser = XMLReaderFactory.createXMLReader();
			SheetHandler handler = new SheetHandler(r.getSharedStringsTable(), reader);
			handler.setStartRow(this.startRow);
			handler.setEndRowIgnoredNum(this.endRowIgnoredNum);
			parser.setContentHandler(handler);
			
			for (int i=1; ; i++) {
				try (InputStream sheet = r.getSheet("rId" + i)) {
					InputSource sheetSource = new InputSource(sheet);
					parser.parse(sheetSource);
				} catch (IllegalArgumentException e) {
					break;
				}
			}
		} catch (Exception e) {
			throw new ExcelReadException("Sax模式读取excel发生异常", e);
		}
	}
	
	private static class SheetHandler extends DefaultHandler {
		
		private SharedStringsTable sst;
		private RowReader reader;
		
		private String lastContents;
		private boolean nextIsString;
		private List<String> row;
		
		private int startRow;
		private int endRowIgnoredNum;
		private int lastRowNum;

		private boolean ignoreRow;
		private String currentColRef;
		
		private SheetHandler(SharedStringsTable sst, RowReader reader) {
			this.sst = sst;
			this.reader = reader;
			this.row = new ArrayList<>();
		}
		
		public void setStartRow(int startRow) {
			this.startRow = startRow;
		}

		public void setEndRowIgnoredNum(int endRowIgnoredNum) {
			this.endRowIgnoredNum = endRowIgnoredNum;
		}
		
		public void startElement(String uri, String localName, String name,
				Attributes attributes) throws SAXException {

			if (name.equals("dimension")) {
				String dimension = attributes.getValue("ref"); //获取数值区域，内容形如：A1:H10

				int separateIndex = dimension.indexOf(":");

				String endCell = dimension.substring(separateIndex + 1);

				Pattern pattern= Pattern.compile("(\\d+)");
				Matcher matcher = pattern.matcher(endCell);

				if (matcher.find()) {
					String endCellRowNo = matcher.group(1).toString();
					this.lastRowNum = Integer.parseInt(endCellRowNo);
				}
			}

			if (name.equals("row")) {
				int currentRowNum = Integer.parseInt(attributes.getValue("r"));

				if (startRow > 1 && startRow > currentRowNum) {
					ignoreRow = true;
				} else {
					ignoreRow = false;
				}

				if (!ignoreRow) {
					if (currentRowNum > (lastRowNum - endRowIgnoredNum)) {
						ignoreRow = true;
					}
				}
			}
			
			// c => cell
			if(name.equals("c")) {
				// Print the cell reference
				// System.out.print(attributes.getValue("r") + " - ");
				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				if(cellType != null && cellType.equals("s")) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
				String newColRef = attributes.getValue("r");
				if (currentColRef == null) {
					currentColRef = newColRef;
				}
			    coverColumnDistanceWithNulls(currentColRef, newColRef);
			    
			    currentColRef = newColRef;
			}
			// Clear contents cache
			lastContents = "";
		}
		
		public void endElement(String uri, String localName, String name)
				throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if(nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				nextIsString = false;
			}

			// v => contents of a cell
			// Output after we've seen the string contents
			if(name.equals("v")) {
				this.row.add(lastContents);
			}
			
			if (name.equals("row")) {
				if (!ignoreRow)
					reader.readRow(row);
				row.clear();
			}
		}
		
		@Override
		public void characters(char[] ch, int start, int length) throws SAXException {
			lastContents += new String(ch, start, length);
		}
		
		private void coverColumnDistanceWithNulls(String fromColRefString, String toColRefString) {
		    int colRefDistance = getDistance(fromColRefString, toColRefString);
		    while (colRefDistance > 1) {
		    	this.row.add(null);
		        --colRefDistance;
		    }
		}

		private int getDistance(String fromColRefString, String toColRefString) {
		    String fromColRef = getExcelCellRef(fromColRefString);
		    String toColRef = getExcelCellRef(toColRefString);
		    int distance = 0;
		    if (fromColRef == null || fromColRef.compareTo(toColRef) > 0)
		        return getDistance("A", toColRefString) + 1;
		    if (fromColRef != null && toColRef != null) {
		        while (fromColRef.length() < toColRef.length() || fromColRef.compareTo(toColRef) < 0) {
		            distance++;
		            fromColRef = increment(fromColRef);
		        }
		    }
		    return distance;
		}
		
		private String increment(String s) {
		    int length = s.length();
		    char c = s.charAt(length - 1);

		    if(c == 'Z') {
		        return length > 1 ? increment(s.substring(0, length - 1)) + 'A' : "AA";
		    }

		    return s.substring(0, length - 1) + ++c;
		}

		private String getExcelCellRef(String fromColRef) {
		    if (fromColRef != null) {
		        int i = 0;
		        for (;i < fromColRef.length(); i++) {
		            if (Character.isDigit(fromColRef.charAt(i))) {
		                break;
		            }
		        }
		        if (i == 0) {
		            return fromColRef;
		        }
		        else {
		            return fromColRef.substring(0, i);
		        }
		    }
		    return null;
		}
	}

	@Override
	public void close() throws IOException{
		if(opcPackage != null) {
			opcPackage.close();
		}
	}
}
