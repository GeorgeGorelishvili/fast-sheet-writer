package org.george.gorelishvili.export.writer;

import org.george.gorelishvili.export.common.ColumnHeader;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.List;

public interface FastSheetWriter {

	void addRow() throws IOException;

	void endRow() throws IOException;

	void addEmptyRow() throws IOException;

	void addCell(Object value) throws IOException;

	void addCell(Object value, String styleKey) throws IOException;

	void addFormulaCell(String formula) throws IOException;

	void addFormulaCell(String formula, String styleKey) throws IOException;

	void addFormulaCell(int columnIndex, String formula, String styleKey) throws IOException;

	void mergeCellsHorizontal(int rowIndex, int firstColumn, int lastColumn) throws IOException;

	int getCurrentRowIndex();

	int getCurrentColumnIndex();

	void createSheet() throws IOException;

	void createSheet(String sheetName) throws IOException;

	void finishSheet() throws IOException;

	void addStyle(String key, XSSFCellStyle style);

	void saveReport(String reportPath, String fileName) throws IOException;

	public XSSFWorkbook getWorkbook();

	void addFirstCell(Object value) throws IOException;

	void addFirstCell(Object value, String styleKey) throws IOException;

	void addNewRowCell(Object value) throws IOException;

	void addNewRowCell(Object value, String styleKey) throws IOException;

	void addLastCell(Object value) throws IOException;

	void addLastCell(Object value, String styleKey) throws IOException;

	void mergeCellsHorizontal(int firstColumn, int lastColumn) throws IOException;

	void addData(List<ColumnHeader> headers, List<Object[]> data) throws IOException;

	void setStartRowIndex(int startRowIndex);

	void setStartColumnIndex(int startColumnIndex);
}
