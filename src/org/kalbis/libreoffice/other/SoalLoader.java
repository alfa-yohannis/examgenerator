package org.kalbis.libreoffice.other;

import java.io.File;
import java.io.IOException;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.container.XIndexContainer;
import com.sun.star.container.XNameAccess;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.TextContentAnchorType;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextFieldsSupplier;
import com.sun.star.text.XTextFrame;
import com.sun.star.text.XTextFramesSupplier;
import com.sun.star.text.XTextGraphicObjectsSupplier;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XWordCursor;
import com.sun.star.uno.Any;
import com.sun.star.uno.Exception;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;

public class SoalLoader {

	public static void main(String[] args) {

		String soalPath = "D:\\LibreOffice\\template.odt";

		try {
			XComponentContext xContext = null;

			// get the remote office component context
			xContext = Bootstrap.bootstrap();
			System.out.println("Connected to a running office ...");

			// get the remote office service manager
			XMultiComponentFactory xMCF = xContext.getServiceManager();

			Object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);

			XComponentLoader xCompLoader = UnoRuntime.queryInterface(XComponentLoader.class, oDesktop);

			// get the file path and transform it into the correct form
			String sUrl = soalPath;
			if (sUrl.indexOf("private:") != 0) {
				File sourceFile = new File(soalPath);
				StringBuffer sbTmp = new StringBuffer("file:///");
				sbTmp.append(sourceFile.getCanonicalPath().replace('\\', '/'));
				sUrl = sbTmp.toString();
			}

			// Load a Writer document, which will be automatically displayed
			XComponent xComp = xCompLoader.loadComponentFromURL(sUrl, "_blank", 0, new PropertyValue[0]);

//			XTextDocument xTextDocument = UnoRuntime.queryInterface(XTextDocument.class, xComp);

			XTextFramesSupplier xTextFramesSupplier = UnoRuntime.queryInterface(XTextFramesSupplier.class, xComp);
			XTextFrame xTextFrame = UnoRuntime.queryInterface(XTextFrame.class,
					xTextFramesSupplier.getTextFrames().getByName("GambarAtauKode"));
//			System.out.println("GambarAtauKode:\r\n " + xTextFrame.getText().getString());

			// -------------
			XTextDocument xTextDocument = UnoRuntime.queryInterface(XTextDocument.class, xComp);

			com.sun.star.lang.XMultiServiceFactory xDocMSF = UnoRuntime
					.queryInterface(com.sun.star.lang.XMultiServiceFactory.class, xTextDocument);
			XInterface obj = (XInterface) xDocMSF.createInstance("com.sun.star.text.TextGraphicObject");
			XTextContent content = UnoRuntime.queryInterface(XTextContent.class, obj);
			
			
			XPropertySet properties = UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, obj);
			properties.setPropertyValue("GraphicURL", "file:///D:/LibreOffice/kalbis.png");
			
			XText xText = xTextDocument.getText();
			XTextCursor xTCursor = xText.createTextCursor();
			xText.insertTextContent(xTCursor, content, false);
			
			// -------------
			XTextGraphicObjectsSupplier xTextGraphicObjectsSupplier = UnoRuntime
					.queryInterface(XTextGraphicObjectsSupplier.class, xComp);

			XTextContent xTextContent = UnoRuntime.queryInterface(XTextContent.class,
					xTextGraphicObjectsSupplier.getGraphicObjects().getByName("Gambar1"));
//			XShape x = UnoRuntime.queryInterface(com.sun.star.drawing.XShape.class, xTextContent);

			String textFrameString = xTextFrame.getText().getString();
			String imageString = xTextContent.getAnchor().getText().getString();

			System.out.println(textFrameString);
			System.out.println("---------------");
			System.out.println(imageString);
			System.out.println("---------------");
			if (textFrameString.equals(imageString)) {
				System.out.println("The image is inside the text frame!");
			} else {
				System.out.println("The image is NOT inside the text frame!");
			}

			xComp.dispose();
			System.exit(0);

			System.console();
		} catch (Exception | BootstrapException | IOException e) {
			System.err.println(" Exception " + e);
			e.printStackTrace(System.err);
			System.exit(0);
		}
	}
}
