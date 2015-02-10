package org.george.gorelishvili.export.common;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public interface SheetWriter {

	void addRow(int row) throws IOException;

	void addEmptyRow(int row) throws IOException;

	void addCell(int columnIndex, Object value) throws IOException;

	void addCell(int columnIndex, Object value, String styleKey) throws IOException;

	void addFormulaCell(int columnIndex, String formula) throws IOException;

	void addFormulaCell(int columnIndex, String formula, String styleKey) throws IOException;

	void createDefaultStyles();

	void createSheet(String sheetName) throws IOException;

	void beginSheet() throws IOException;

	void endSheet() throws IOException;

	void endRow() throws IOException;

	void addStyle(String key, XSSFCellStyle style);

	void mergeCellsHorizontal(int row, int firstColumn, int lastColumn);

	void closeWriter() throws IOException;

	XSSFWorkbook getWorkbook();

	void saveReport(String reportPath, String fileName) throws IOException;
}
