package org.kalbis.libreoffice.other;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLDocumentPart.RelationPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.wp.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFEndnote;
import org.apache.poi.xwpf.usermodel.XWPFFieldRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

public class TestPoi {

	public static void main(String[] args) {
		try {
			File file = new File("D:\\LibreOffice\\template.docx");
			FileInputStream fis = new FileInputStream(file.getAbsolutePath());
			XWPFDocument document = new XWPFDocument(fis);
			
//			for (POIXMLDocumentPart x : document.getRelations()) {
//				PackagePart y = x.getPackagePart();
//				System.out.println(y.getPartName());
//				System.out.println(x.toString());
//				System.console();
//			}
			
			for (XWPFTable t : document.getTables()) {
				for (XWPFTableRow r : t.getRows()) {
					for (XWPFTableCell c : r.getTableCells()) {
						c.getParagraphs().get(0).getCTP();
						System.out.print(c.getText());
						System.out.print(" ");
					}
					System.out.println();
				}
				System.console();
			}
			
//			for (XWPFPictureData p : document.getAllPictures()) {
//				Object a1 = p.getData();
//				Object a2 = p.getFileName();
//				Object a3 = p.getPackagePart();
//				Object a4 = p.getParent();
//				Object a5 = p.getPictureType();
//				Object a6 = p.getRelationParts();
//				Object a7 = p.getRelations();
//				Object a8 = p.getPackagePart().getPackage();
//				Object a9 = p.getPackagePart().getPartName();
//				System.console();
//			}
//			
//			for (XWPFPictureData p : document.getAllPackagePictures()) {
//				Object a1 = p.getData();
//				Object a2 = p.getFileName();
//				Object a3 = p.getPackagePart();
//				Object a4 = p.getParent();
//				Object a5 = p.getPictureType();
//				Object a6 = p.getRelationParts();
//				Object a7 = p.getRelations();
//				Object a8 = p.getPackagePart().getPackage();
//				Object a9 = p.getPackagePart().getPartName();
//				System.console();
//			}

//			for (XWPFParagraph paragraph : document.getParagraphs()) {
//				XmlCursor cursor = paragraph.getCTP().newCursor();
//				cursor.selectPath(
//						"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:drawing/*/w:txbxContent/w:p/w:r");
//
//				List<XmlObject> ctrsintxtbx = new ArrayList<XmlObject>();
//
//				while (cursor.hasNextSelection()) {
//					cursor.toNextSelection();
//					XmlObject obj = cursor.getObject();
////					System.out.println(obj.toString());
//
//					CTR ctr = CTR.Factory.parse(obj.xmlText());
//					XWPFRun bufferrun = new XWPFRun(ctr, (IRunBody) paragraph);
//					System.out.println(bufferrun.text());
//				}
////				for (XmlObject obj : ctrsintxtbx) {
////					CTR ctr = CTR.Factory.parse(obj.xmlText());
////					// CTR ctr = CTR.Factory.parse(obj.newInputStream());
////					XWPFRun bufferrun = new XWPFRun(ctr, (IRunBody) paragraph);
////					String text = bufferrun.getText(0);
////					System.out.println(text);
//////					if (text != null && text.contains(someWords)) {
//////						text = text.replace(someWords, "replaced");
//////						bufferrun.setText(text, 0);
//////					}
//////					obj.set(bufferrun.getCTR());
////				}
//			}

			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
