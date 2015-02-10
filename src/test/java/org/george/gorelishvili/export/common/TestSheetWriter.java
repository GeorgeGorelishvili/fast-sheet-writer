package org.george.gorelishvili.export.common;

import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;

public class TestSheetWriter {

	private static final String REPORT_PATH = "c:/tmp";
	private static final String FILE_NAME = "test";

	@Test
	public void normalizeFileNameTest() throws IOException {
		SheetWriterImpl writer = new SheetWriterImpl();
		String filename = "first";
		Assert.assertEquals(writer.normalizeFileName(filename), filename + ".xlsx");

		filename = "second.xlsx";
		Assert.assertEquals(writer.normalizeFileName(filename), filename);
	}
}
