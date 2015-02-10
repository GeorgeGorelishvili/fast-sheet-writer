package org.george.gorelishvili.export.common;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class SheetWriterImpl implements SheetWriter {

	private XSSFWorkbook workbook;
	private Writer writer;
	private int rowIndex;
	Map<String, XSSFCellStyle> styles = new HashMap<>();
	private Set<String> mergeCells = new HashSet<>();

	private static final String XLSX = ".xlsx";
	private static final String REPORT_FILENAME = "report";
	private static final String FILE_ENCODING = "UTF-8";
	private static final String XML = ".xml";

	private File template;
	private File sheet;
	private String referenceName;

	public SheetWriterImpl() {
		workbook = new XSSFWorkbook();
	}

	@Override
	public void createSheet(String sheetName) throws IOException {
		XSSFSheet worksheet = workbook.createSheet(sheetName);
		referenceName = worksheet.getPackagePart().getPartName().getName().substring(1);

		template = File.createTempFile(REPORT_FILENAME, XLSX);
		FileOutputStream fos = new FileOutputStream(template);
		workbook.write(fos);
		fos.close();

		sheet = File.createTempFile(sheetName, XML);
		writer = new OutputStreamWriter(new FileOutputStream(sheet), FILE_ENCODING);
		beginSheet();
	}

	@Override
	public void beginSheet() throws IOException {
		writer.write("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n" +
				"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" );
		writer.write("<sheetData>\n");
	}

	@Override
	public void endSheet() throws IOException{
		writer.write("</sheetData>");
		appendMergeCells();
		writer.write("</worksheet>");
	}

	@Override
	public void addRow(int row) throws IOException {
		writer.write(String.format("<row r=\"%d\">\n", row + 1));
		this.rowIndex = row;
	}

	@Override
	public void addEmptyRow(int row) throws IOException {
		addRow(row);
		endRow();
	}

	@Override
	public void endRow() throws IOException {
		writer.write("</row>\n");
	}

	@Override
	public void addCell(int columnIndex, Object value) throws IOException {
		addCell(columnIndex, value, null);
	}

	@Override
	public void addCell(int columnIndex, Object value, String styleKey) throws IOException {
		addCell(columnIndex, value, styleKey, false);
	}

	@Override
	public void addFormulaCell(int columnIndex, String formula) throws IOException {
		addFormulaCell(columnIndex, formula, null);
	}

	@Override
	public void addFormulaCell(int columnIndex, String formula, String styleKey) throws IOException {
		addCell(columnIndex, formula, styleKey, true);
	}

	@Override
	public void addStyle(String key, XSSFCellStyle style) {
		styles.put(key, style);
	}

	@Override
	public void closeWriter() throws IOException {
		writer.close();
	}

	@Override
	public void mergeCellsHorizontal(int row, int firstColumn, int lastColumn) {
		mergeCells.add(new CellRangeAddress(row, row, firstColumn, lastColumn).formatAsString());
	}

	private void addCell(int columnIndex, Object value, String styleKey, boolean formula) throws IOException {
		String cellReference = new CellReference(this.rowIndex, columnIndex).formatAsString();
		writer.write(new CellBuilder.Builder(cellReference)
				.style(getStyle(styleKey))
				.value(value)
				.formula(formula)
				.create());
	}

	private void appendMergeCells() throws IOException {
		if (!mergeCells.isEmpty()) {
			writer.write("<mergeCells count=\"");
			writer.write(String.valueOf(mergeCells.size()));
			writer.write("\">");
			for (String ref : mergeCells) {
				writer.write("<mergeCell ref=\"");
				writer.write(ref);
				writer.write("\"/>");
			}
			writer.write("</mergeCells>");
		}
	}

	String normalizeFileName(String fileName) {
		if (fileName == null) {
			return REPORT_FILENAME + XLSX;
		}
		String result = fileName.trim();
		if (!result.endsWith(XLSX)) {
			result = fileName + XLSX;
		}
		return result;
	}

	String normalizeDirectory(String dir) {
		if (dir == null) {
			return "c:/";
		}
		String result = dir.trim();
		return result.endsWith("/") ? result.substring(0, result.length() - 1) : result;
	}

	@Override
	public void saveReport(String reportDirPath, String fileName) throws IOException {
		fileName = normalizeFileName(fileName);
		reportDirPath = normalizeDirectory(reportDirPath);
		File dir = new File(reportDirPath);
		if (!dir.exists() && !dir.mkdir()) {
			throw new IOException("Could not create report directory!");
		}
		String filePath = reportDirPath + "/" + fileName;
		createFile(filePath);
	}

	@SuppressWarnings("unchecked")
	private void createFile(String filePath) throws IOException {
		ZipFile zip = new ZipFile(template);
		FileOutputStream out = new FileOutputStream(filePath);
		ZipOutputStream zos = new ZipOutputStream(out);

		Enumeration<ZipEntry> en = (Enumeration<ZipEntry>) zip.entries();
		while (en.hasMoreElements()) {
			ZipEntry ze = en.nextElement();
			if(!ze.getName().equals(referenceName)){
				zos.putNextEntry(new ZipEntry(ze.getName()));
				InputStream is = zip.getInputStream(ze);
				copyStream(is, zos);
				is.close();
			}
		}

		zos.putNextEntry(new ZipEntry(referenceName));
		InputStream is = new FileInputStream(sheet);
		copyStream(is, zos);
		is.close();

		zos.close();
		zip.close();
	}

	@Override
	public void createDefaultStyles() {

		XSSFDataFormat fmt = workbook.createDataFormat();

		XSSFCellStyle style1 = workbook.createCellStyle();
		style1.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
		style1.setDataFormat(fmt.getFormat("0.00"));
		addStyle(Keys.AMOUNT, style1);

		XSSFCellStyle style2 = workbook.createCellStyle();
		style2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		style2.setDataFormat(fmt.getFormat("dd/MM/yyyy"));
		addStyle(Keys.DATE, style2);

		XSSFCellStyle style3 = workbook.createCellStyle();
		style3.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		style3.setDataFormat(fmt.getFormat("dd/MM/yyyy hh:mm"));
		addStyle(Keys.DATETIME, style3);

		XSSFCellStyle newLineStyle = workbook.createCellStyle();
		newLineStyle.setWrapText(true);
		addStyle(Keys.ALLOW_WRAP, newLineStyle);
	}

	@Override
	public XSSFWorkbook getWorkbook() {
		return workbook;
	}

	private void copyStream(InputStream in, OutputStream out) throws IOException {
		byte[] buffer = new byte[1024];
		int read;
		while ((read = in.read(buffer)) >= 0) {
			out.write(buffer, 0, read);
		}
	}

	private XSSFCellStyle getStyle(String styleKey) {
		return styleKey != null ? styles.get(styleKey) : null;
	}
}
