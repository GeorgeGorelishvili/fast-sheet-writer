package org.george.gorelishvili.export.common;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.junit.Assert;
import org.junit.Test;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class TestCellBuilder {

	private static final String cellReference = "A1";

	private static final String	STRING_VALUE = "value1";
	private static final Integer INTEGER_VALUE = 15;
	private static final Long LONG_VALUE = 85l;
	private static final Short SHORT_VALUE = 45;
	private static final Double AMOUNT_VALUE = 15.45;
	private static final Boolean BOOLEAN_VALUE = true;
	private static final String FORMULA_VALUE = "SUM(A1:A5)";

	@Test
	public void nullValueTest() {
		String result = new CellBuilder.Builder(cellReference)
				.value(null)
				.formula(false)
				.create();
		Assert.assertEquals(result, "");
	}

	@Test
	public void dateValueTest() throws ParseException{
		SimpleDateFormat df = new SimpleDateFormat(Format.DATE);
		String value = "20/11/2014";
		Date DATE_VALUE = df.parse(value);
		String result = new CellBuilder.Builder(cellReference)
				.style(null)
				.value(DATE_VALUE)
				.create();
		Assert.assertEquals(result, getDateResult(null, value));


		SheetWriterImpl sw = new SheetWriterImpl();
		sw.createDefaultStyles();
		XSSFCellStyle style = sw.styles.get(Keys.DATETIME);

		df = new SimpleDateFormat(Format.DATETIME);
		value = "10/11/2014 03:27";
		DATE_VALUE = df.parse(value);
		result = new CellBuilder.Builder(cellReference)
				.value(DATE_VALUE)
				.style(style)
				.create();
		Assert.assertEquals(result, getDateResult(style, value));
	}

	@Test
	public void shortValueTest() {
		String result = new CellBuilder.Builder(cellReference)
				.style(null)
				.value(SHORT_VALUE)
				.create();
		Assert.assertEquals(result, getNumberResult(null, SHORT_VALUE));
	}

	@Test
	public void integerValueTest() {
		String result = new CellBuilder.Builder(cellReference)
				.style(null)
				.value(INTEGER_VALUE)
				.create();
		Assert.assertEquals(result, getNumberResult(null, INTEGER_VALUE));
	}

	@Test
	public void stringValueTest() {
		String result = new CellBuilder.Builder(cellReference)
				.style(null)
				.value(STRING_VALUE)
				.create();
		Assert.assertEquals(result, getStringResult(null));

		SheetWriterImpl sw = new SheetWriterImpl();
		sw.createDefaultStyles();
		XSSFCellStyle style = sw.styles.get(Keys.ALLOW_WRAP);

		result = new CellBuilder.Builder(cellReference)
				.value(STRING_VALUE)
				.style(style)
				.create();
		Assert.assertEquals(result, getStringResult(style));
	}

	@Test
	public void formulaValueTest() {
		String result = new CellBuilder.Builder(cellReference)
				.style(null)
				.value(FORMULA_VALUE)
				.formula(true)
				.create();
		Assert.assertEquals(result, getFormulaResult(null));

		SheetWriterImpl sw = new SheetWriterImpl();
		sw.createDefaultStyles();
		XSSFCellStyle style = sw.styles.get(Keys.ALLOW_WRAP);

		result = new CellBuilder.Builder(cellReference)
				.value(STRING_VALUE)
				.style(style)
				.create();
		Assert.assertEquals(result, getStringResult(style));
	}

	@Test
	public void doubleValueTest() {
		String result = new CellBuilder.Builder(cellReference)
				.style(null)
				.value(AMOUNT_VALUE)
				.create();
		Assert.assertEquals(result, getDoubleResult(null));

		SheetWriterImpl sw = new SheetWriterImpl();
		sw.createDefaultStyles();
		XSSFCellStyle style = sw.styles.get(Keys.AMOUNT);

		result = new CellBuilder.Builder(cellReference)
				.value(AMOUNT_VALUE)
				.style(style)
				.create();
		Assert.assertEquals(result, getDoubleResult(style));
	}

	@Test
	public void booleanValueTest() {
		String result = new CellBuilder.Builder(cellReference)
				.style(null)
				.value(BOOLEAN_VALUE)
				.create();

		Assert.assertEquals(result, getBooleanValue(null, BOOLEAN_VALUE ? 1 : 0));
	}

	@Test
	public void longValueTest() {
		String result = new CellBuilder.Builder(cellReference)
				.style(null)
				.value(LONG_VALUE)
				.create();
		Assert.assertEquals(result, getNumberResult(null, LONG_VALUE));
	}

	private String getStringResult(XSSFCellStyle style) {
		StringBuilder b = new StringBuilder();
		b.append("<c");
		b.append(" r=\"").append(cellReference).append("\"");
		b.append(" t=\"inlineStr\"");
		if (style != null) {
			b.append(" s=\"").append(style.getIndex()).append("\"");
		}
		b.append(">");
		b.append("<is><t>").append(STRING_VALUE).append("</t></is>");
		b.append("</c>\n");
		return b.toString();
	}

	private String getFormulaResult(XSSFCellStyle style) {
		StringBuilder b = new StringBuilder();
		b.append("<c");
		b.append(" r=\"").append(cellReference).append("\"");
		b.append(" t=\"inlineStr\"");
		if (style != null) {
			b.append(" s=\"").append(style.getIndex()).append("\"");
		}
		b.append(">");
		b.append("<f>").append(FORMULA_VALUE).append("</f>");
		b.append("</c>\n");
		return b.toString();
	}

	private String getDateResult(XSSFCellStyle style, String date) {
		StringBuilder b = new StringBuilder();
		b.append("<c");
		b.append(" r=\"").append(cellReference).append("\"");
		b.append(" t=\"inlineStr\"");
		if (style != null) {
			b.append(" s=\"").append(style.getIndex()).append("\"");
		}
		b.append(">");
		b.append("<is><t>").append(date).append("</t></is>");
		b.append("</c>\n");
		return b.toString();
	}

	private String getDoubleResult(XSSFCellStyle style) {
		StringBuilder b = new StringBuilder();
		b.append("<c");
		b.append(" r=\"").append(cellReference).append("\"");
		b.append(" t=\"n\"");
		if (style != null) {
			b.append(" s=\"").append(style.getIndex()).append("\"");
		}
		b.append(">");
		b.append("<v>").append(AMOUNT_VALUE).append("</v>");
		b.append("</c>\n");
		return b.toString();
	}

	private String getNumberResult(XSSFCellStyle style, Number number) {
		StringBuilder b = new StringBuilder();
		b.append("<c");
		b.append(" r=\"").append(cellReference).append("\"");
		if (style != null) {
			b.append(" s=\"").append(style.getIndex()).append("\"");
		}
		b.append(">");
		b.append("<v>").append(number).append("</v>");
		b.append("</c>\n");
		return b.toString();
	}

	private String getBooleanValue(XSSFCellStyle style, int value) {
		StringBuilder b = new StringBuilder();
		b.append("<c");
		b.append(" r=\"").append(cellReference).append("\"");
//		b.append(" t=\"n\"");
		if (style != null) {
			b.append(" s=\"").append(style.getIndex()).append("\"");
		}
		b.append(">");
		b.append("<v>").append(value).append("</v>");
		b.append("</c>\n");
		return b.toString();
	}
}
