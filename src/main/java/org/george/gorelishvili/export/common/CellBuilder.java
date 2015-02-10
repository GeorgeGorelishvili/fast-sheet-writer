package org.george.gorelishvili.export.common;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.text.SimpleDateFormat;
import java.util.Date;

public class CellBuilder {

	private CellBuilder() {}

	public static class Builder {
		private String cellReference;
		private XSSFCellStyle style;
		private Object value;
		private boolean formula;

		public Builder(String cellReference) {
			this.cellReference = cellReference;
		}

		public Builder style(XSSFCellStyle style) {
			this.style = style;
			return this;
		}

		public Builder value(Object value) {
			this.value = value;
			return this;
		}

		public Builder formula(boolean formula) {
			this.formula = formula;
			return this;
		}

		public String create() {
			StringBuilder builder = new StringBuilder();
			if (value != null) {
				builder.append("<c r=\"").append(cellReference).append("\"");
				if (value instanceof String) {
					builder.append(" t=\"inlineStr\"");
				} else if (value instanceof Double) {
					builder.append(" t=\"n\"");
				} else if (value instanceof Date) {
					builder.append(" t=\"inlineStr\"");
				}
				if (style != null) {
					builder.append(" s=\"").append(style.getIndex()).append("\"");
				}
				builder.append(">");
				if (value instanceof Date) {
					String format = "dd/MM/yyyy";
					if (style != null && style.getDataFormatString() != null) {
						format = style.getDataFormatString();
					}
					SimpleDateFormat df = new SimpleDateFormat(format);
					builder.append("<is><t>").append(df.format(value)).append("</t></is>");
				} else if (value instanceof Boolean) {
					builder.append("<v>").append((((Boolean) value) ? 1 : 0)).append("</v>");
				} else if (value instanceof String) {
					if (formula) {
						builder.append("<f>").append(value).append("</f>");
					} else {
						builder.append("<is><t>").append(value).append("</t></is>");
					}
				} else {
					builder.append("<v>").append(value).append("</v>");
				}
				builder.append("</c>\n");
			}
			return builder.toString();
		}
	}
}
