package org.kalbis.libreoffice;

import static org.junit.Assert.*;

public class Test {

	@org.junit.Test
	public void testRandomCode() {
		String code = ExamPerStudentSession.generateCode(6);
		System.out.println(code);
		assertEquals(true, Integer.parseInt(code) >= 0);
		assertEquals(6, code.length());
	}

}
