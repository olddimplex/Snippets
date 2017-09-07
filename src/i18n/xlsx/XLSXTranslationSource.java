package i18n.xlsx;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.Map;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipInputStream;

import javax.swing.JFileChooser;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import common.IStringTransformer;

/**
 * Parses an .xlsx file or a classpath resource to extract string translation mappings.
 */
public class XLSXTranslationSource {

	private class WorksheetHandlerMap extends DefaultHandler {

		private final Map<LngKey, String> translationMap;
		private final IStringTransformer transformer;
		private final StringBuilder sb = new StringBuilder();
		
		private final ArrayList<String> path = new ArrayList<String>();
		private final String rowPath[] = new String[]{"worksheet","sheetData","row"};
		private final String columnPath[] = new String[]{"worksheet","sheetData","row","c"};
		private final String valuePath[] = new String[]{"worksheet","sheetData","row","c","v"};

		private int columnIndex = 0;
		private String columnType;
		private String languageCode;
		private String translationKey;

		public WorksheetHandlerMap(final Map<LngKey, String> out, final IStringTransformer transformer) {
			this.translationMap = out;
			this.transformer = (transformer == null) ? IStringTransformer.ECHO: transformer;
		}

		@Override
		public void startElement(final String uri, final String localName, final String qName, final Attributes attributes) throws SAXException {
			final WorksheetHandlerMap self = WorksheetHandlerMap.this;
			self.path.add((localName.isEmpty()) ? qName : localName);
			if(Arrays.equals(self.path.toArray(), self.rowPath)) {
				self.columnIndex = 0;
			} else
			if(Arrays.equals(self.path.toArray(), self.columnPath) && self.columnIndex < 3) {
				self.columnType = attributes.getValue("t");
			}

		}
		
		@Override
		public void endElement(final String uri, final String localName, final String qName) throws SAXException {
			final WorksheetHandlerMap self = WorksheetHandlerMap.this;
			if(Arrays.equals(self.path.toArray(), self.columnPath)) {
				switch(self.columnIndex) {
					case 0: 
						self.languageCode = self.columnValue(); 
						break;
					case 1: 
						self.translationKey = self.columnValue(); 
						break;
					case 2: 
						self.translationMap.put(new LngKey(self.languageCode, self.translationKey), self.columnValue()); 
						break;
					default: 
						break;
				}
				self.columnIndex++;
				self.sb.setLength(0);
			}
			self.path.remove(self.path.size() - 1);
		}

		@Override
		public void characters(final char[] ch, final int start, final int length) throws SAXException {
			final WorksheetHandlerMap self = WorksheetHandlerMap.this;
			if(Arrays.equals(self.path.toArray(), self.valuePath) && self.columnIndex < 3) try {
				self.sb.append(ch, start, length);
			} catch(Exception e) {
				throw new SAXException(e);
			}
		}
		
		private String columnValue() {
			final WorksheetHandlerMap self = WorksheetHandlerMap.this;
			final XLSXTranslationSource parent = XLSXTranslationSource.this;
			if("s".equals(self.columnType)) {
				return transformer.transform(parent.sharedStrings.get(Integer.parseInt(self.sb.toString())));
			} else {
				return transformer.transform(self.sb.toString());
			}
		}
	}

	private class WorksheetHandlerFile extends DefaultHandler {

		private final Writer writer;
		private final StringBuilder sb = new StringBuilder();
		
		private final ArrayList<String> path = new ArrayList<String>();
		private final String rowPath[] = new String[]{"worksheet","sheetData","row"};
		private final String columnPath[] = new String[]{"worksheet","sheetData","row","c"};
		private final String valuePath[] = new String[]{"worksheet","sheetData","row","c","v"};

		private final String rowDelimiter;
		private String currentRowDelimiter = "";
		private final String columnDelimiter;
		private String currentColumnDelimiter = "";
		private int columnIndex = 0;
		private String columnType;

		public WorksheetHandlerFile(final Writer out, final String columnDelimiter, final String rowDelimiter) throws UnsupportedEncodingException {
			this.writer = out;
			this.columnDelimiter = columnDelimiter;
			this.rowDelimiter = rowDelimiter;
		}

		@Override
		public void startElement(final String uri, final String localName, final String qName, final Attributes attributes) throws SAXException {
			final WorksheetHandlerFile self = WorksheetHandlerFile.this;
			self.path.add((localName.isEmpty()) ? qName : localName);
			try {
				if(Arrays.equals(self.path.toArray(), self.rowPath)) {
					self.writer.write(self.currentRowDelimiter);
					self.columnIndex = 0;
					self.currentColumnDelimiter = "";
					self.currentRowDelimiter = self.rowDelimiter;
				} else
				if(Arrays.equals(self.path.toArray(), self.columnPath) && self.columnIndex < 3) {
					self.columnType = attributes.getValue("t");
					self.writer.write(self.currentColumnDelimiter);
					self.currentColumnDelimiter = self.columnDelimiter;
				}
			} catch(Exception e) {
				throw new SAXException(e);
			}
		}
		
		@Override
		public void endElement(final String uri, final String localName, final String qName) throws SAXException {
			final WorksheetHandlerFile self = WorksheetHandlerFile.this;
			final XLSXTranslationSource parent = XLSXTranslationSource.this;
			if(Arrays.equals(self.path.toArray(), self.columnPath)) {
				if(self.columnIndex < 3) try {
					if("s".equals(self.columnType)) {
						self.writer.write(parent.sharedStrings.get(Integer.parseInt(self.sb.toString())));
					} else {
						self.writer.write(self.sb.toString());
					}
				} catch(final Exception e) {
					throw new SAXException(e);
				}
				self.columnIndex++;
				self.sb.setLength(0);
			}
			self.path.remove(self.path.size() - 1);
		}

		@Override
		public void characters(final char[] ch, final int start, final int length) throws SAXException {
			final WorksheetHandlerFile self = WorksheetHandlerFile.this;
			if(Arrays.equals(self.path.toArray(), self.valuePath) && self.columnIndex < 3) try {
				self.sb.append(ch, start, length);
			} catch(Exception e) {
				throw new SAXException(e);
			}
		}
	}
	
	private class SharedStringsHandler extends DefaultHandler {

		private final StringBuilder sb = new StringBuilder();
		
		private final ArrayList<String> path = new ArrayList<String>();
		private final String targetPath[] = new String[]{"sst","si","t"};
		
		@Override
		public void startElement(final String uri, final String localName, final String qName, final Attributes attributes) throws SAXException {
			final SharedStringsHandler self = SharedStringsHandler.this;
			self.path.add((localName.isEmpty()) ? qName : localName);
			if(Arrays.equals(self.path.toArray(), self.targetPath)) {
				self.sb.setLength(0);
			}
		}
		
		@Override
		public void endElement(final String uri, final String localName, final String qName) throws SAXException {
			final SharedStringsHandler self = SharedStringsHandler.this;
			final XLSXTranslationSource parent = XLSXTranslationSource.this;
			if(Arrays.equals(self.path.toArray(), self.targetPath)) {
				parent.sharedStrings.add(self.sb.toString());
			}
			self.path.remove(self.path.size() - 1);
		}
		
		@Override
		public void characters(final char[] ch, final int start, final int length) throws SAXException {
			final SharedStringsHandler self = SharedStringsHandler.this;
			if(Arrays.equals(self.path.toArray(), self.targetPath)) {
				self.sb.append(ch, start, length);
			}
		}
	}
	
	private static final Pattern SHARED_STRINGS_PATTERN = Pattern.compile("xl/sharedStrings\\.xml");
	private static final Pattern WORKSHEET_PATTERN = Pattern.compile("xl/worksheets/[^/]+\\.xml");
	
	private final ArrayList<String> sharedStrings = new ArrayList<String>();
	private final ArrayList<ZipEntry> worksheetEntries = new ArrayList<ZipEntry>();
	
	private final String resourcePath;
	private final File inputFile;

	public static void main(final String[] args) throws Exception {
		final JFileChooser fileChooser = new JFileChooser();
		fileChooser.setDialogTitle("Pick the input .xlsx file");
		if(JFileChooser.APPROVE_OPTION == fileChooser.showOpenDialog(null)) {
			final XLSXTranslationSource xlsxProcessor = new XLSXTranslationSource(fileChooser.getSelectedFile());
//			final ConcurrentHashMap<LngKey,String> map = new ConcurrentHashMap<LngKey, String>();
//			xlsxProcessor.export(map, new UnquotingStringTransformer());
//			for(final Entry<LngKey,String> entry : map.entrySet()) {
//				System.out.printf("%s,%s,%s\n", entry.getKey().getLanguage(),entry.getKey().getPhrase(),entry.getValue());
//			}
			fileChooser.setDialogTitle("Pick the output CSV file");
			if(JFileChooser.APPROVE_OPTION == fileChooser.showOpenDialog(null)) {
				try(final Writer outWriter = new UnquotingWriter(new FileOutputStream(fileChooser.getSelectedFile()), "UTF-16")) {
					xlsxProcessor.export(outWriter, ",", "\n");
				}
			}
		}
	}

	/**
	 * Examines if the given classpath resource is an XLSX file with the expected structure.<br/>
	 * Each worksheet is expected to have three columns, where:<br/>
	 * - the first column contains the language code;<br/>
	 * - the second column contains the key phrase;<br/>
	 * - the third column contains the translated phrase.
	 * 
	 * @param resourcePath
	 * @throws SAXException
	 * @throws IOException
	 * @throws ParserConfigurationException
	 */
	public XLSXTranslationSource(final String resourcePath) throws SAXException, IOException, ParserConfigurationException {
		this.resourcePath = resourcePath;
		this.inputFile = null;
		try (final ZipInputStream zin = new ZipInputStream(Thread.currentThread().getContextClassLoader().getResourceAsStream(resourcePath))) {
			for(ZipEntry entry; (entry = zin.getNextEntry()) != null && !entry.isDirectory();) {
				this.zipEntry(zin, entry);
			}
		}
		this.validate();
	}

	/**
	 * Examines if the given file is an XLSX file with the expected structure.<br/>
	 * Each worksheet is expected to have three columns, where:<br/>
	 * - the first column contains the language code;<br/>
	 * - the second column contains the key phrase;<br/>
	 * - the third column contains the translated phrase.
	 * 
	 * @param xlsxFile
	 * @throws SAXException
	 * @throws IOException
	 * @throws ParserConfigurationException
	 */
	public XLSXTranslationSource(final File xlsxFile) throws SAXException, IOException, ParserConfigurationException  {
		this.resourcePath = null;
		this.inputFile = xlsxFile;
		try (final ZipFile zipFile = new ZipFile(this.inputFile)) {
			final Enumeration<? extends ZipEntry> zipEntries = zipFile.entries();
			while(zipEntries.hasMoreElements()) {
				final ZipEntry entry = (ZipEntry)zipEntries.nextElement();
				if(!entry.isDirectory()) {
					this.zipEntry(zipFile.getInputStream(entry), entry);
				}
			}
		}
		this.validate();
	}
	
	private void validate() {
		if(this.worksheetEntries.isEmpty()) {
			throw new IllegalArgumentException("No Excel worksheets found.");
		}
	}

	private void zipEntry(final InputStream in, final ZipEntry zipEntry) throws SAXException, IOException, ParserConfigurationException {
		if(SHARED_STRINGS_PATTERN.matcher(zipEntry.getName()).matches()) {
			this.sharedStrings(in);
		} else
		if(WORKSHEET_PATTERN.matcher(zipEntry.getName()).matches()) {
			this.worksheetEntries.add(zipEntry);
		} else {
//			System.out.printf("Name: %s, isDirectory: %s\n", zipEntry.getName(), zipEntry.isDirectory());
		}
	}
	
	private void sharedStrings(final InputStream in) throws SAXException, IOException, ParserConfigurationException  {
		final SharedStringsHandler handler = new SharedStringsHandler();
		// prevent the SAXParser from closing the given InputStream
		SAXParserFactory.newInstance().newSAXParser().parse(new InputStream() {
			@Override
			public int read() throws IOException {
				return in.read();
			}
		}, handler);
	}


	/**
	 * Parses all worksheet data and writes respective records to the given writer.
	 * 
	 * @param writer
	 * @param columnDelimiter
	 * @param rowDelimiter
	 * @throws SAXException
	 * @throws IOException
	 * @throws ParserConfigurationException
	 */
	public void export(
			  final Writer writer
			, final String columnDelimiter
			, final String rowDelimiter
			) throws SAXException, IOException, ParserConfigurationException {
		if(this.inputFile != null) {
			this.exportInputFileToCSV(writer, columnDelimiter, rowDelimiter);
		}
		if(this.resourcePath != null) {
			this.exportClasspathResourceToCSV(writer, columnDelimiter, rowDelimiter);
		}
	}
	
	private void exportInputFileToCSV(
			  final Writer writer
			, final String columnDelimiter
			, final String rowDelimiter
			) throws SAXException, IOException, ParserConfigurationException {
		try(final ZipFile zipFile = new ZipFile(this.inputFile)) {
			final WorksheetHandlerFile handler = new WorksheetHandlerFile(writer, columnDelimiter, rowDelimiter);
			for(final ZipEntry entry : this.worksheetEntries) {
				SAXParserFactory.newInstance().newSAXParser().parse(zipFile.getInputStream(entry), handler);
			}
		}
	}
	
	private void exportClasspathResourceToCSV(
			  final Writer writer
			, final String columnDelimiter
			, final String rowDelimiter
			) throws SAXException, IOException, ParserConfigurationException  {
		try(final InputStream in = Thread.currentThread().getContextClassLoader().getResourceAsStream(this.resourcePath);
			final ZipInputStream zin = new ZipInputStream(in)) {
			final WorksheetHandlerFile handler = new WorksheetHandlerFile(writer, columnDelimiter, rowDelimiter);
			for(ZipEntry entry; (entry = zin.getNextEntry()) != null && !entry.isDirectory();) {
				if(WORKSHEET_PATTERN.matcher(entry.getName()).matches()) {
					// prevent the SAXParser from closing the ZipInputStream
					SAXParserFactory.newInstance().newSAXParser().parse(new InputStream() {
						@Override
						public int read() throws IOException {
							return zin.read();
						}
					}, handler);
				}
			}
		}
	}

	/**
	 * Parses all worksheet data and adds respective entries into the given map.
	 * 
	 * @param translations
	 * @throws SAXException
	 * @throws IOException
	 * @throws ParserConfigurationException
	 */
	public void export(final Map<LngKey,String> translations) throws SAXException, IOException, ParserConfigurationException {
		if(this.inputFile != null) {
			this.exportInputFileToMap(translations, IStringTransformer.ECHO);
		}
		if(this.resourcePath != null) {
			this.exportClasspathResourceToMap(translations, IStringTransformer.ECHO);
		}
	}

	/**
	 * Parses all worksheet data and adds respective entries into the given map.<br/>
	 * The {@link LngKey#LngKey(String, String)} arguments, as well as respective entry values,
	 * will be passed to the given transformer right before use.
	 * 
	 * @param translations
	 * @param transformer
	 * @throws SAXException
	 * @throws IOException
	 * @throws ParserConfigurationException
	 */
	public void export(final Map<LngKey,String> translations, final IStringTransformer transformer) throws SAXException, IOException, ParserConfigurationException {
		if(this.inputFile != null) {
			this.exportInputFileToMap(translations, transformer);
		}
		if(this.resourcePath != null) {
			this.exportClasspathResourceToMap(translations, transformer);
		}
	}

	private void exportInputFileToMap(final Map<LngKey,String> translations, final IStringTransformer transformer) throws SAXException, IOException, ParserConfigurationException {
		try (final ZipFile zipFile = new ZipFile(this.inputFile)) {
			final WorksheetHandlerMap handler = new WorksheetHandlerMap(translations, transformer);
			for(final ZipEntry entry : this.worksheetEntries) {
				SAXParserFactory.newInstance().newSAXParser().parse(zipFile.getInputStream(entry), handler);
			}
		}
	}
	
	private void exportClasspathResourceToMap(final Map<LngKey,String> translations, final IStringTransformer transformer) throws SAXException, IOException, ParserConfigurationException {
		try(final InputStream in = Thread.currentThread().getContextClassLoader().getResourceAsStream(this.resourcePath)) {
			final WorksheetHandlerMap handler = new WorksheetHandlerMap(translations, transformer);
			try (final ZipInputStream zin = new ZipInputStream(in)) {
				for(ZipEntry entry; (entry = zin.getNextEntry()) != null && !entry.isDirectory();) {
					if(WORKSHEET_PATTERN.matcher(entry.getName()).matches()) {
						// prevent the SAXParser from closing the ZipInputStream
						SAXParserFactory.newInstance().newSAXParser().parse(new InputStream() {
							@Override
							public int read() throws IOException {
								return zin.read();
							}
						}, handler);
					}
				}
			}
		}
	}
}
