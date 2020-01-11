package snippets.xlsx;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.function.Consumer;
import java.util.function.Predicate;
import java.util.function.Supplier;
import java.util.function.UnaryOperator;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipInputStream;

import javax.swing.JFileChooser;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * Parses an .xlsx file or a classpath resource to ...
 */
public class XLSXAwareSource<T extends XLSXAware> {

	private static final Logger LOGGER = LogManager.getLogger(XLSXAwareSource.class);

	private class WorksheetHandler extends DefaultHandler {

		private final UnaryOperator<String> transformer;
		private final StringBuilder sb = new StringBuilder();
		
		private final ArrayList<String> path = new ArrayList<String>();
		private final String worksheetPath[] = new String[]{"worksheet"};
		private final String rowPath[] = new String[]{"worksheet","sheetData","row"};
		private final String columnPath[] = new String[]{"worksheet","sheetData","row","c"};
		private final String valuePath[] = new String[]{"worksheet","sheetData","row","c","v"};

		private int worksheetIndex = 0;
		private boolean worksheetProcessingAllowed = false;
		private int rowIndex = 0;
		private boolean rowProcessingAllowed = false;
		private int columnIndex = 0;
		private boolean columnProcessingAllowed = false;
		private String rowName;
		private String columnName;
		private String columnType;
		private T xlsxAware;

		public WorksheetHandler(final UnaryOperator<String> transformer) {
			this.transformer = (transformer == null) ? UnaryOperator.<String> identity(): transformer;
		}

		@Override
		public void startElement(final String uri, final String localName, final String qName, final Attributes attributes) throws SAXException {
			final XLSXAwareSource<T> parent = XLSXAwareSource.this;
			final WorksheetHandler self = WorksheetHandler.this;
			self.path.add((localName.isEmpty()) ? qName : localName);
			if(Arrays.equals(self.path.toArray(), self.worksheetPath)) {
				self.worksheetProcessingAllowed = XLSXAwareSource.this.worksheetSelector.test(self.worksheetIndex++);
				self.rowIndex = 0;
			} else
			if(self.worksheetProcessingAllowed && Arrays.equals(self.path.toArray(), self.rowPath)) {
				self.rowProcessingAllowed = XLSXAwareSource.this.rowSelector.test(self.rowIndex++);
				if(self.rowProcessingAllowed) {
					self.rowName = attributes.getValue("r");
					self.columnIndex = 0;
					self.xlsxAware = parent.supplier.get();
				}
			} else
			if(self.worksheetProcessingAllowed && self.rowProcessingAllowed && Arrays.equals(self.path.toArray(), self.columnPath)) {
				self.columnProcessingAllowed = XLSXAwareSource.this.columnSelector.test(self.columnIndex++);
				if(self.columnProcessingAllowed) {
					self.columnName = attributes.getValue("r");
					self.columnName = self.columnName.substring(0, self.columnName.length() - self.rowName.length());
					self.columnType = attributes.getValue("t");
				}
			}
		}
		
		@Override
		public void endElement(final String uri, final String localName, final String qName) throws SAXException {
			final WorksheetHandler self = WorksheetHandler.this;
			if(self.worksheetProcessingAllowed && self.rowProcessingAllowed && Arrays.equals(self.path.toArray(), self.columnPath)) {
				if(self.columnProcessingAllowed) try {
					self.xlsxAware.set(self.columnName, self.columnValue());
				} catch(final Exception e) {
					LOGGER.error("", e);
				}
				self.sb.setLength(0);
			} else
			if(self.worksheetProcessingAllowed && self.rowProcessingAllowed && Arrays.equals(self.path.toArray(), self.rowPath)) {
				XLSXAwareSource.this.consumer.accept(self.xlsxAware);
			}
			self.path.remove(self.path.size() - 1);
		}

		@Override
		public void characters(final char[] ch, final int start, final int length) throws SAXException {
			final WorksheetHandler self = WorksheetHandler.this;
			if(self.worksheetProcessingAllowed && self.rowProcessingAllowed && self.columnProcessingAllowed && Arrays.equals(self.path.toArray(), self.valuePath)) try {
				self.sb.append(ch, start, length);
			} catch(Exception e) {
				throw new SAXException(e);
			}
		}
		
		private String columnValue() {
			final WorksheetHandler self = WorksheetHandler.this;
			final XLSXAwareSource<T> parent = XLSXAwareSource.this;
			if("s".equals(self.columnType)) {
				return transformer.apply(parent.sharedStrings.get(Integer.parseInt(self.sb.toString())));
			} else {
				return transformer.apply(self.sb.toString());
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
			final XLSXAwareSource<T> parent = XLSXAwareSource.this;
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
	
	private final ArrayList<String> sharedStrings = new ArrayList<>();
	private final ArrayList<ZipEntry> worksheetEntries = new ArrayList<>();
	
	private final String resourcePath;
	private final File inputFile;
	private final Supplier<T> supplier;
	private final Consumer<T> consumer;
	private final Predicate<Integer> worksheetSelector;
	private final Predicate<Integer> rowSelector;
	private final Predicate<Integer> columnSelector;

	public static void main(final String[] args) throws Exception {
		final JFileChooser fileChooser = new JFileChooser();
		fileChooser.setDialogTitle("Pick the input .xlsx file");
		if(JFileChooser.APPROVE_OPTION == fileChooser.showOpenDialog(null)) {
//			final List<ExpertProfileFormAggregation> list = new ArrayList<>();
//			final XLSXExpertProfileAccountSource<ExpertProfileFormAggregation> xlsxProcessor = 
//				new XLSXExpertProfileAccountSource<>(
//					fileChooser.getSelectedFile(), 
//					() -> new ExpertProfileFormAggregation(), 
//					(expertProfileFormAggregation) -> {
//						if(expertProfileFormAggregation.getEmailArray().length > 1) {
//							LOGGER.error(expertProfileFormAggregation.getPersonalNumber() + "\t" + Arrays.toString(expertProfileFormAggregation.getEmailArray()));
//						}
////						if(!ValidationUtil.isValidPNOBG(expertProfileFormAggregation.getPersonalNumber(), LOGGER)) { 
////							list.add(expertProfileFormAggregation);
////						}
//					},
//					(worksheetIndex) -> worksheetIndex == 1, 
//					(columnIndex) -> columnIndex < 20
//				);
//			xlsxProcessor.export();
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
	public XLSXAwareSource(
			final String resourcePath, 
			final Supplier<T> supplier, 
			final Consumer<T> consumer, 
			final Predicate<Integer> worksheetIndexSelector,
			final Predicate<Integer> rowIndexSelector,
			final Predicate<Integer> columnIndexSelector
			) throws SAXException, IOException, ParserConfigurationException {
		this.resourcePath = resourcePath;
		this.supplier = supplier;
		this.consumer = (consumer == null) ? (type) -> {} : consumer;
		this.worksheetSelector = (worksheetIndexSelector == null) ? (worksheetIndex) -> true : worksheetIndexSelector;
		this.rowSelector = (rowIndexSelector == null) ? (rowIndex) -> true : rowIndexSelector;
		this.columnSelector = (columnIndexSelector == null) ? (columnIndex) -> true : columnIndexSelector;
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
	public XLSXAwareSource(
			final File xlsxFile, 
			final Supplier<T> supplier, 
			final Consumer<T> consumer, 
			final Predicate<Integer> worksheetIndexSelector,
			final Predicate<Integer> rowIndexSelector,
			final Predicate<Integer> columnIndexSelector
			) throws SAXException, IOException, ParserConfigurationException  {
		this.resourcePath = null;
		this.supplier = supplier;
		this.consumer = (consumer == null) ? (type) -> {} : consumer;
		this.worksheetSelector = (worksheetIndexSelector == null) ? (worksheetIndex) -> true : worksheetIndexSelector;
		this.rowSelector = (rowIndexSelector == null) ? (rowIndex) -> true : rowIndexSelector;
		this.columnSelector = (columnIndexSelector == null) ? (columnIndex) -> true : columnIndexSelector;
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
	 * Parses all worksheet data and adds respective entries into the given map.
	 * 
	 * @param list
	 * @throws SAXException
	 * @throws IOException
	 * @throws ParserConfigurationException
	 */
	public void export() throws SAXException, IOException, ParserConfigurationException {
		if(this.inputFile != null) {
			this.exportInputFileToList(UnaryOperator.<String> identity());
		}
		if(this.resourcePath != null) {
			this.exportClasspathResourceToList(UnaryOperator.<String> identity());
		}
	}

	/**
	 * Parses all worksheet data and adds respective entries into the given map.<br/>
	 * The {@link LngKey#LngKey(String, String)} arguments, as well as respective entry values,
	 * will be passed to the given transformer right before use.
	 * 
	 * @param list
	 * @param transformer
	 * @throws SAXException
	 * @throws IOException
	 * @throws ParserConfigurationException
	 */
	public void export(final UnaryOperator<String> transformer) throws SAXException, IOException, ParserConfigurationException {
		if(this.inputFile != null) {
			this.exportInputFileToList(transformer);
		}
		if(this.resourcePath != null) {
			this.exportClasspathResourceToList(transformer);
		}
	}

	private void exportInputFileToList(final UnaryOperator<String> transformer) throws SAXException, IOException, ParserConfigurationException {
		try (final ZipFile zipFile = new ZipFile(this.inputFile)) {
			final WorksheetHandler handler = new WorksheetHandler(transformer);
			for(final ZipEntry entry : this.worksheetEntries) {
				SAXParserFactory.newInstance().newSAXParser().parse(zipFile.getInputStream(entry), handler);
			}
		}
	}
	
	private void exportClasspathResourceToList(final UnaryOperator<String> transformer) throws SAXException, IOException, ParserConfigurationException {
		try(final InputStream in = Thread.currentThread().getContextClassLoader().getResourceAsStream(this.resourcePath)) {
			final WorksheetHandler handler = new WorksheetHandler(transformer);
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
