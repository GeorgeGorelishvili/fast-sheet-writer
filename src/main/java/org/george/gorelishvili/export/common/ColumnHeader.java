package org.george.gorelishvili.export.common;

public class ColumnHeader {

	private String name;

	private String styleKey;

	private ColumnHeader() {}

	private ColumnHeader(String name) {
		this.name = name;
	}

	private ColumnHeader(String name, String styleKey) {
		this.name = name;
		this.styleKey = styleKey;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getStyleKey() {
		return styleKey;
	}

	public void setStyleKey(String styleKey) {
		this.styleKey = styleKey;
	}
}
