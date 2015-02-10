package org.george.gorelishvili.export.writer;

import org.george.gorelishvili.export.common.SheetWriterImpl;
import org.george.gorelishvili.export.common.ColumnHeader;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.List;

public class FastSheetWriterImpl implements FastSheetWriter {

	private static final String SHEET_NAME = "sheet";
	private SheetWriterImpl sw;

	private int START_ROW_INDEX;
	private int START_COLUMN_INDEX;
	private int rowIndex;
	private int columnIndex;

	boolean isRowOpened;
	boolean isSheetCreated;

	public static FastSheetWriter getInstance() {
		return new FastSheetWriterImpl();
	}

	FastSheetWriterImpl() {
		sw = new SheetWriterImpl();
		sw.createDefaultStyles();
	}

	@Override
	public void createSheet() throws IOException {
		createSheet(null);
	}

	@Override
	public void createSheet(String name) throws IOException {
		isSheetCreated = true;
		sw.createSheet(name != null ? name : SHEET_NAME);
	}

	@Override
	public void finishSheet() throws IOException {
		endRowIfNeed();
		isSheetCreated = false;
		sw.endSheet();
		sw.closeWriter();
	}

	@Override
	public void addRow() throws IOException {
		createSheetIfNeed();
		endRowIfNeed();
		isRowOpened = true;
		sw.addRow(rowIndex);
	}

	@Override
	public void addEmptyRow() throws IOException {
		endRowIfNeed();
		endRow();
	}

	@Override
	public void endRow() throws IOException {
		addNewRowIfNeed();
		isRowOpened = false;
		columnIndex = START_COLUMN_INDEX;
		rowIndex++;
		sw.endRow();
	}

	@Override
	public void addStyle(String key, XSSFCellStyle style) {
		sw.addStyle(key, style);
	}

	@Override
	public void addCell(Object value) throws IOException {
		addCell(value, null);
	}

	@Override
	public void addCell(Object value, String styleKey) throws IOException {
		addNewRowIfNeed();
		sw.addCell(columnIndex, value, styleKey);
		columnIndex++;
	}

	@Override
	public void addFormulaCell(String formula) throws IOException {
		addFormulaCell(formula, null);
	}

	@Override
	public void addFormulaCell(String formula, String styleKey) throws IOException {
		addNewRowIfNeed();
		sw.addFormulaCell(columnIndex, formula, styleKey);
		columnIndex++;
	}

	@Override
	public void addFormulaCell(int columnIndex, String formula, String styleKey) throws IOException {
		this.columnIndex = columnIndex + START_COLUMN_INDEX;
		addFormulaCell(formula, styleKey);
	}

	@Override
	public void mergeCellsHorizontal(int firstColumn, int lastColumn) throws IOException {
		mergeCellsHorizontal(rowIndex, firstColumn, lastColumn);
	}

	@Override
	public void mergeCellsHorizontal(int rowIndex, int firstColumn, int lastColumn) throws IOException {
		addNewRowIfNeed();
		sw.mergeCellsHorizontal(rowIndex, firstColumn, lastColumn);
	}

	@Override
	public void addFirstCell(Object value) throws IOException {
		addFirstCell(value, null);
	}

	@Override
	public void addFirstCell(Object value, String styleKey) throws IOException {
		endRow();
		addCell(value, styleKey);
	}

	@Override
	public void addNewRowCell(Object value) throws IOException {
		addNewRowCell(value, null);
	}

	@Override
	public void addNewRowCell(Object value, String styleKey) throws IOException {
		addCell(value, styleKey);
	}

	@Override
	public void addLastCell(Object value) throws IOException {
		addLastCell(value, null);
	}

	@Override
	public void addLastCell(Object value, String styleKey) throws IOException {
		addCell(value, styleKey);
		endRow();
	}

	@Override
	public XSSFWorkbook getWorkbook() {
		return sw.getWorkbook();
	}

	@Override
	public void saveReport(String reportPath, String fileName) throws IOException {
		finishSheetIfNeed();
		sw.saveReport(reportPath, fileName);
	}

	@Override
	public void addData(List<ColumnHeader> headers, List<Object[]> data) throws IOException {
		for (ColumnHeader header : headers) {
			addCell(header.getName());
		}
		endRow();

		for (Object[] objects : data) {
			for (int i = 0; i < objects.length; i++) {
				addCell(objects[i], headers.get(i).getStyleKey());
			}
			endRow();
		}
	}

	@Override
	public int getCurrentRowIndex() {
		return rowIndex;
	}

	@Override
	public int getCurrentColumnIndex() {
		return columnIndex + START_COLUMN_INDEX;
	}

	@Override
	public void setStartRowIndex(int startRowIndex) {
		START_ROW_INDEX = startRowIndex > 0 ? startRowIndex : 0;
		this.rowIndex = START_ROW_INDEX;

	}

	@Override
	public void setStartColumnIndex(int startColumnIndex) {
		START_COLUMN_INDEX = startColumnIndex > 0 ? startColumnIndex : 0;
		this.columnIndex = START_COLUMN_INDEX;
	}

	private void addNewRowIfNeed() throws IOException {
		if (!isRowOpened) {
			addRow();
		}
	}

	private void endRowIfNeed() throws IOException {
		if (isRowOpened) {
			endRow();
		}
	}

	private void createSheetIfNeed() throws IOException {
		if (!isSheetCreated) {
			createSheet();
		}
	}

	private void finishSheetIfNeed() throws IOException {
		if (isSheetCreated) {
			finishSheet();
		}
	}
}
