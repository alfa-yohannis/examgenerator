package org.kalbis.libreoffice;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Random;

public class Question {

	public static final Random RANDOM = new Random();
	private String code;
	private int session;

	private String description;
	private List<Option> options;
	private List<Option> randomisedOptions;
	private Option answer;

	public Question(String code, int session, String description, List<Option> options, Option answer) {
		this.code = code;
		this.session = session;
		this.description = description;
		this.options = options;
		this.answer = answer;
		this.randomisedOptions = new ArrayList<Option>();
	}

	public int getAnswerIndexOfRandomisedOptions() {
		return this.randomisedOptions.indexOf(this.answer);
	}
	
	public void randomiseOptions() {
		this.randomisedOptions.clear();
		List<Integer> order = new ArrayList<>(Arrays.asList(new Integer[] { 0, 1, 2, 3, 4 }));
		while (order.size() > 0) {
			int temp = RANDOM.nextInt(order.size());
			int index = order.get(temp);
			this.randomisedOptions.add(this.options.get(index));
			order.remove(temp);
		}
	}

	public List<Option> getOptions() {
		return options;
	}

	public String getCode() {
		return code;
	}

	public int getSession() {
		return session;
	}

	public String getDescription() {
		return description;
	}

	public Option getAnswer() {
		return answer;
	}

	public List<Option> getRandomiseOptions() {
		return randomisedOptions;
	}

}
