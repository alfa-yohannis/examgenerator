package org.kalbis.libreoffice;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Random;

public class ExamPerStudentSession {
	private String code;
	private Student student;

	private List<Question> questions;
	private List<Question> randomisedQuestions;

	public ExamPerStudentSession(Student student) {
		code = generateCode(6);
		this.student = student;
		questions = new ArrayList<Question>();
		randomisedQuestions = new ArrayList<Question>();
	}

	public ExamPerStudentSession() {
		this(null);
	}

	public void randomiseQuestions() {
		this.randomisedQuestions.clear();
		List<Integer> order = new ArrayList<Integer>();
		for (int i = 0; i < questions.size(); i++) {
			order.add(i);
		}
		while (order.size() > 0) {
			int temp = Question.RANDOM.nextInt(order.size());
			int index = order.get(temp);
			this.randomisedQuestions.add(this.questions.get(index));
			order.remove(temp);
		}
	}

	public static String generateCode(int length) {
		Random random = new Random();
		String code = "";
		for (int i = 0; i < length; i++) {
			code = code + String.valueOf(random.nextInt(10));
		}
		return code;
	}

	public Student getStudent() {
		return student;
	}

	public String getCode() {
		return code;
	}

	public void setStudent(Student student) {
		this.student = student;
	}

	public List<Question> getRandomisedQuestions() {
		return randomisedQuestions;
	}

	public List<Question> getQuestions() {
		return questions;
	}

}
