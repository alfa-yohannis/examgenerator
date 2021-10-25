package org.kalbis.examgenerator;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.CopyOption;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ExamQuestionsGenerator {

	private static final String TEMPLATE_SOAL_DOCX = "template-soal-UTS.docx";
	private static final String TEORI_LEMBAR_JAWABAN = "uts_teori_lembar_jawaban";
	private static final String TEORI_JAWABAN = "uts_teori_jawaban";
	private static final String TEORI_SOAL = "uts_teori_soal";
	private static final String STUDENTS_XLSX = "students.xlsx";
	private static final String BANK_SOAL_ANDROID_XLSX = "bank_soal_android-UTS.xlsx";
	private static final String LEMBAR_JAWABAN_XLSX = "lembar-jawaban-UTS.xlsx";
	private static final String workingDir ="D:\\KALBIS\\mobile-21-22-ganjil\\";

	public static void main(String[] args) throws Exception {
		
		// load questions from question bank
		File questionBankFile = new File(workingDir + BANK_SOAL_ANDROID_XLSX);

		// empty target directory both questions, answer keys, and lembar jawaban
		File questionsTargetDir = new File(workingDir + TEORI_SOAL);
		if (!questionsTargetDir.exists())
			questionsTargetDir.mkdir();
		for (File file : questionsTargetDir.listFiles())
			if (!file.isDirectory())
				file.delete();

		File answerKeyTargetDir = new File(workingDir + TEORI_JAWABAN);
		if (!answerKeyTargetDir.exists())
			answerKeyTargetDir.mkdir();
		for (File file : answerKeyTargetDir.listFiles())
			if (!file.isDirectory())
				file.delete();

		File answerSheetTargetDir = new File(workingDir + TEORI_LEMBAR_JAWABAN);
		if (!answerSheetTargetDir.exists())
			answerSheetTargetDir.mkdir();
		for (File file : answerSheetTargetDir.listFiles())
			if (!file.isDirectory())
				file.delete();

		// read students data
		File studentListFile = new File(workingDir + STUDENTS_XLSX);
		List<Student> students = loadStudentList(studentListFile);

		// edit and save the generated document one by one
		for (Student student : students) {
			ExamPerStudentSession examPerStudentSession = createExamPerStudentSession(student, questionBankFile);
			updateDocumentAndCreateAnswerKeys(examPerStudentSession, questionsTargetDir, answerKeyTargetDir,
					answerSheetTargetDir);
		}

		System.out.println("Finished!");
	}

	private static void updateDocumentAndCreateAnswerKeys(ExamPerStudentSession examPerStudentSession,
			File questionsTargetDir, File answerKeyTargetDir, File answerSheetTargetDir) throws IOException {

		Student student = examPerStudentSession.getStudent();
		System.out.print("Generating documents for " + student.getCode() + " " + student.getName() + " ... ");

		examPerStudentSession.randomiseQuestions();

		// copy template to target dir
		String templatePath = workingDir + TEMPLATE_SOAL_DOCX;
		String newName = questionsTargetDir.getAbsolutePath() + File.separator + student.getCode() + "_"
				+ student.getName() + "_" + examPerStudentSession.getCode() + ".docx";
		Files.copy(Paths.get(templatePath), Paths.get(newName), StandardCopyOption.REPLACE_EXISTING);
		File examFile = new File(newName);

		// create answer keys file in csv
		String answerKeyFileName = answerKeyTargetDir.getAbsolutePath() + File.separator + student.getCode() + "_"
				+ student.getName() + "_" + examPerStudentSession.getCode() + ".csv";
		File answerKeysFile = new File(answerKeyFileName);
		FileWriter fw = new FileWriter(answerKeysFile);
		BufferedWriter bw = new BufferedWriter(fw);

		// create header
		bw.append("student,exam,no,question,answer");
		bw.newLine();

		// open docx file for editing
		FileInputStream fis = new FileInputStream(examFile.getAbsolutePath());
		XWPFDocument document = new XWPFDocument(fis);

		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setBold(true);
		run.setFontSize(14);
		run.setText("Kode Soal: " + examPerStudentSession.getCode());
		run.addBreak();

		run = paragraph.createRun();
		run.setText("NIM: " + examPerStudentSession.getStudent().getCode());
		run.addBreak();
		run.setText("Nama: " + examPerStudentSession.getStudent().getName());
		run.addBreak();

		for (int i = 0; i < examPerStudentSession.getRandomisedQuestions().size(); i++) {
			Question question = examPerStudentSession.getRandomisedQuestions().get(i);

			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setBold(true);
			run.setText("" + (i + 1));
			run.setText(". ");

			run = paragraph.createRun();
			run.setBold(false);
			run.setText(question.getDescription());
			run.addBreak();
			run = paragraph.createRun();
			run.setBold(true);
			run.setText("Jawaban:");
			run.addBreak();

			question.randomiseOptions();
			for (int j = 0; j < question.getRandomiseOptions().size(); j++) {
				Option option = question.getRandomiseOptions().get(j);
				run = paragraph.createRun();
				run.setBold(true);
				if (j == 0) {
					run.setText("A. ");
				} else if (j == 1) {
					run.setText("B. ");
				} else if (j == 2) {
					run.setText("C. ");
				} else if (j == 3) {
					run.setText("D. ");
				} else if (j == 4) {
					run.setText("E. ");
				}
				run = paragraph.createRun();
				run.setText(option.getDescription());
				run.addBreak();
			}
			int answerIndex = question.getAnswerIndexOfRandomisedOptions();
			String answerKey = null;
			switch (answerIndex) {
			case 0:
				answerKey = "A";
				break;
			case 1:
				answerKey = "B";
				break;
			case 2:
				answerKey = "C";
				break;
			case 3:
				answerKey = "D";
				break;
			case 4:
				answerKey = "E";
				break;
			default:
				break;
			}

			// student,exam,no,question,answer
			String line = student.getCode() + "," + examPerStudentSession.getCode() + "," + (i + 1) + ","
					+ question.getCode() + "," + answerKey;
			bw.append(line);
			bw.newLine();
		}

		OutputStream outputStream = new FileOutputStream(examFile.getAbsolutePath());
		document.write(outputStream);
		outputStream.close();
		document.close();
		bw.close();
		fw.close();
		fis.close();

		// create an answer sheet for each student
		String answerSheetTemplatePath = workingDir + File.separator + LEMBAR_JAWABAN_XLSX;
		String studentAnswerSheet = answerSheetTargetDir.getAbsolutePath() + File.separator + student.getCode() + "_"
				+ student.getName() + ".xlsx";
		Files.copy(Paths.get(answerSheetTemplatePath), Paths.get(studentAnswerSheet),
				StandardCopyOption.REPLACE_EXISTING);

		FileInputStream inputFile = new FileInputStream(studentAnswerSheet);
		XSSFWorkbook workbook = new XSSFWorkbook(inputFile);
		XSSFSheet sheet = workbook.getSheetAt(0);

//		sheet.getRow(0).getCell(1).setCellValue(examPerStudentSession.getCode());
		sheet.getRow(7).getCell(1).setCellValue(student.getCode());
		sheet.getRow(8).getCell(1).setCellValue(student.getName());

		// update the properties as well
		POIXMLProperties props = workbook.getProperties();
		POIXMLProperties.CoreProperties coreProp = props.getCoreProperties();
		coreProp.setCreator("Alfa Yohannis"); // set document creator
		coreProp.setTitle(student.getCode() + "_" + student.getName());
		coreProp.setSubjectProperty(
				student.getCode() + "_" + student.getName());
		coreProp.setDescription(student.getCode() + "_" + student.getName());

		FileOutputStream outputFile = new FileOutputStream(studentAnswerSheet);
		workbook.write(outputFile);
		workbook.close();
		inputFile.close();

		System.out.println(" finished");
	}

	public static List<Student> loadStudentList(File studentListFile) throws IOException {
		List<Student> students = new ArrayList<Student>();
		FileInputStream fis = new FileInputStream(studentListFile);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);

		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			String code = String.valueOf((int) row.getCell(1).getNumericCellValue());
			String name = row.getCell(2).getStringCellValue();
			Student student = new Student(code, name);
//			System.out.println(name);
			students.add(student);
		}

		workbook.close();
		fis.close();
		return students;
	}

	/***
	 * Construct the question set from the question bank.
	 * 
	 * @param questionBankFile
	 * @param answerKeyTargetDir
	 * @return
	 * @throws Exception
	 */

	public static ExamPerStudentSession createExamPerStudentSession(Student student, File questionBankFile)
			throws Exception {
		ExamPerStudentSession examPerStudentSession = new ExamPerStudentSession(student);
		Question question = null;
		String questionCode = null;
		int questionSession = -1;
		String questionDescription = null;
		List<Option> options = null;
		String answerString = null;

		FileInputStream fis = new FileInputStream(questionBankFile);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);

		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			// skip header
			if (row.getRowNum() == 0) {
				continue;
			} else if (row.getRowNum() % 5 == 1) {
				questionCode = new String(row.getCell(0).getStringCellValue());
				questionSession = new Integer((int) row.getCell(1).getNumericCellValue());
				questionDescription = new String(row.getCell(2).getStringCellValue());
				options = new ArrayList<Option>();
				answerString = new String(row.getCell(5).getStringCellValue());

			}

			String optionCode = row.getCell(3).getStringCellValue();
			String optionDescription = row.getCell(4).getStringCellValue();
			Option option = new Option(optionCode, optionDescription);
			options.add(option);

			// since every question has five rows (five options)
			// if the row number is the multiple of 5 then create the question
			if (row.getRowNum() % 5 == 0) {
				final String a = answerString;
				Option answer = options.stream().filter(o -> o.getCode().equals(a)).findFirst().orElse(null);
				if (answer == null)
					throw new Exception("No aswer defined!");
				question = new Question(questionCode, questionSession, questionDescription, options, answer);
				examPerStudentSession.getQuestions().add(question);
			}
		}
		workbook.close();
		fis.close();
		return examPerStudentSession;
	}

}
