package org.george.gorelishvili.export.writer;

import org.george.gorelishvili.export.common.Keys;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.IOException;

public class TestFastWriter {
	private static final String REPORT_PATH = System.getProperty("user.dir") + "/tmp";
	private static final String FILE_NAME = "first";

	@Test
	public void fileSaveTest() throws IOException {
		FastSheetWriter fsw = FastSheetWriterImpl.getInstance();
		fsw.mergeCellsHorizontal(1, 2);
		fsw.addCell("first");
		fsw.addCell("second");
		fsw.addFirstCell("third");

		fsw.mergeCells(2, 3, 3, 6);
		fsw.saveReport("/home/george/temp/"/*REPORT_PATH*/, FILE_NAME);
		String filePath = "/home/george/temp/"/*REPORT_PATH*/ + "/" + FILE_NAME + Keys.XLSX;
		Assert.assertTrue(new File(filePath).exists());
	}

	@Test
	public void createSheetTest() throws IOException {
		FastSheetWriterImpl bean = new FastSheetWriterImpl();
		Assert.assertNotNull(bean.getWorkbook());
		bean.createSheet(null);
		Assert.assertEquals(bean.getCurrentRowIndex(), 0);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 0);
		Assert.assertEquals(bean.isRowOpened, false);

		bean.addRow();
		Assert.assertEquals(bean.getCurrentRowIndex(), 0);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 0);
		Assert.assertEquals(bean.isRowOpened, true);

		bean.addRow();
		Assert.assertEquals(bean.getCurrentRowIndex(), 1);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 0);
		Assert.assertEquals(bean.isRowOpened, true);

		bean.addCell(null);
		Assert.assertEquals(bean.getCurrentRowIndex(), 1);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 1);
		Assert.assertEquals(bean.isRowOpened, true);

		bean.addLastCell(null);
		Assert.assertEquals(bean.getCurrentRowIndex(), 2);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 0);
		Assert.assertEquals(bean.isRowOpened, false);

		bean.addFirstCell(null);
		Assert.assertEquals(bean.getCurrentRowIndex(), 3);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 1);
		Assert.assertEquals(bean.isRowOpened, true);

		bean.endRow();
		Assert.assertEquals(bean.getCurrentRowIndex(), 4);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 0);
		Assert.assertEquals(bean.isRowOpened, false);

		bean.endRow();
		Assert.assertEquals(bean.getCurrentRowIndex(), 5);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 0);
		Assert.assertEquals(bean.isRowOpened, false);

		bean.addEmptyRow();
		Assert.assertEquals(bean.getCurrentRowIndex(), 6);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 0);
		Assert.assertEquals(bean.isRowOpened, false);

		bean.addFormulaCell(null);
		Assert.assertEquals(bean.getCurrentRowIndex(), 6);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 1);
		Assert.assertEquals(bean.isRowOpened, true);

		bean.addEmptyRow();
		bean.addNewRowCell(null);
		Assert.assertEquals(bean.getCurrentRowIndex(), 8);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 1);
		Assert.assertEquals(bean.isRowOpened, true);

		bean.finishSheet();
		Assert.assertEquals(bean.getCurrentRowIndex(), 9);
		Assert.assertEquals(bean.getCurrentColumnIndex(), 0);
		Assert.assertEquals(bean.isRowOpened, false);
	}
}
