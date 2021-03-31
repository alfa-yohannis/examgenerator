package org.kalbis.libreoffice;

public class Option {

	private String code;
	private String description;
	
	public Option(String code, String description) {
		this.code = code;
		this.description = description;
	}

	public String getCode() {
		return code;
	}

	public String getDescription() {
		return description;
	}
}
