package org.kalbis.examgenerator;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collector;
import java.util.stream.Collectors;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExamMarksGenerator {

	public static void main(String[] args) throws Exception {

		String workingDirString = "C:\\kalbis\\mobile-20-21-genap\\";
		String hasilDirString = workingDirString + "uas_teori_hasil";
		String answerKeyDirString = workingDirString + "uas_teori_jawaban_202106081554";
		String nilaiDirString = workingDirString + "uas_teori_nilai";
		String recapFileName = "recap.csv";

		File hasilDir = new File(hasilDirString);
		File answerKeyDir = new File(answerKeyDirString);

		List<String> recapLines = new ArrayList<>();

		for (File file : hasilDir.listFiles()) {
//			System.out.println(file.getAbsolutePath());

			FileInputStream inputFile = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(inputFile);
			XSSFSheet sheet = workbook.getSheetAt(0);

			String questionSheetNumber = sheet.getRow(0).getCell(1).getRawValue();
			String studentCode = sheet.getRow(7).getCell(1).getStringCellValue().trim();
			String studentName = sheet.getRow(8).getCell(1).getStringCellValue().trim();

			System.out.print("Kode Soal: " + questionSheetNumber + ", NIM: " + studentCode + ", Nama: " + studentName);

			// find the right answer key file
			File rightKeyFile = null;
			for (File keyFile : answerKeyDir.listFiles()) {
				String temp = keyFile.getName().substring(0, 10);
				if (temp.equals(studentCode)) {
					rightKeyFile = keyFile;
					break;
				}
			}

			if (rightKeyFile == null) {
				System.out.println(" --> File does not exist");
				continue;
			}
			System.out.println();

			FileReader fr = new FileReader(rightKeyFile);
			BufferedReader br = new BufferedReader(fr);
			String line = null;

			int countCorrectAnswer = 0;
			int countWrongAnswer = 0;
			int numOfQuestions = 0;

			String studentFileName = nilaiDirString + File.separator + studentName + "-" + studentCode + ".csv";
			File studentFile = new File(studentFileName);
			FileWriter fw1 = new FileWriter(studentFile);
			BufferedWriter bw1 = new BufferedWriter(fw1);
			bw1.append("sheet_code,no,correct_answer,student_answer,result");
			bw1.newLine();
			// ---
			int lineNum = 1;
			while ((line = br.readLine()) != null) {
				// skip header
				if (lineNum == 1) {
					lineNum++;
					continue;
				}

				String[] parts = line.split(",");
				int questionNum = Integer.valueOf(parts[2]);
				String rightAnswer = parts[4];
				questionSheetNumber = parts[1];

				int padding = 10;
				int rowNum = padding + questionNum;

				String studentAnswer = sheet.getRow(rowNum).getCell(1).getStringCellValue().trim();
				if (studentAnswer.length() > 1) {
					studentAnswer = studentAnswer.substring(0, 1);
				}
				String result;
				if (rightAnswer.toUpperCase().equals(studentAnswer.toUpperCase())) {
					countCorrectAnswer++;
//					System.out.println("TRUE");
					result = "TRUE";
				} else {
					countWrongAnswer++;
//					System.out.println("FALSE");
					result = "FALSE";
				}
				bw1.append(questionSheetNumber + "," + questionNum + "," + rightAnswer + "," + studentAnswer + "," + result);
				bw1.newLine();

				lineNum++;
				numOfQuestions++;
			}
			bw1.close();
			fw1.close();

			double mark = ((double) countCorrectAnswer) / (double) numOfQuestions * 100d;
			System.out.println("Benar = " + countCorrectAnswer + ", Salah = " + countWrongAnswer + ", Nilai = "
					+ String.format("%.2f", mark));
			recapLines.add(studentName + "," + studentCode + "," + questionSheetNumber + "," + numOfQuestions + ","
					+ countCorrectAnswer + "," + countWrongAnswer + "," + String.format("%.2f", mark));

			workbook.close();
		}

		File recapFile = new File(nilaiDirString + File.separator + recapFileName);
		if (recapFile.exists())
			recapFile.delete();
		recapFile.createNewFile();
		FileWriter fw = new FileWriter(recapFile);
		BufferedWriter bw = new BufferedWriter(fw);
		recapLines = recapLines.stream().sorted().collect(Collectors.toList());
		recapLines.add(0,"nim,nama,kode_soal,jumlah_soal,benar,salah,nilai");
		for (String recapLine : recapLines) {
			bw.append(recapLine);
			bw.newLine();
		}
		bw.close();
		fw.close();

		System.out.println("Finished!!!");
	}

}
