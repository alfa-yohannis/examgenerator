package org.kalbis.examgenerator;

public class Student {

	private String code;
	private String name;
	
	public Student(String code, String name) {
		this.code = code;
		this.name = name;
	}

	public String getName() {
		return name;
	}

	public String getCode() {
		return code;
	}

}
