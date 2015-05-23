package convertors;

import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class Convertor {

	public static void main(String[] args) throws IOException, ParserConfigurationException, TransformerException {
		// convertDocxToHTML();
		long st = System.currentTimeMillis();
//		convertDocToHTML("D:/Traditional.doc", "D:/Traditional.html");
		//convertPDFToHTML("D:/Muthu Krishnan T Resume.pdf", "D:/Muthu Krishnan T Resume.html");
		long et = System.currentTimeMillis();
		System.out.println("Time taken: " + (et - st));
	}

	public static void convertDocxToHTML(String docxPath, String htmlPath) throws FileNotFoundException, IOException {
		String fileInName = "test";
		String root = "target";
		String fileOutName = root + "/" + fileInName + ".html";

		long startTime = System.currentTimeMillis();
		InputStream fileIn = new FileInputStream(new File("D:/Traditional.doc"));
		XWPFDocument document = new XWPFDocument(fileIn);

		XHTMLOptions options = XHTMLOptions.create();// .indent( 4 );
		// Extract image
		File imageFolder = new File(root + "/images/" + fileInName);
		options.setExtractor(new FileImageExtractor(imageFolder));
		// URI resolver
		options.URIResolver(new FileURIResolver(imageFolder));

		OutputStream out = new FileOutputStream(new File(fileOutName));
		XHTMLConverter.getInstance().convert(document, out, options);

		System.out.println("Generate " + fileOutName + " with " + (System.currentTimeMillis() - startTime) + " ms.");
	}

	public static void convertDocToHTML(String docPath, String htmlPath) throws ParserConfigurationException, FileNotFoundException, IOException, TransformerException {
		HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(new FileInputStream(docPath));

		WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
		wordToHtmlConverter.processDocument(wordDocument);
		org.w3c.dom.Document htmlDocument = wordToHtmlConverter.getDocument();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		DOMSource domSource = new DOMSource(htmlDocument);
		StreamResult streamResult = new StreamResult(out);

		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer serializer = tf.newTransformer();
		serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
		serializer.setOutputProperty(OutputKeys.INDENT, "yes");
		serializer.setOutputProperty(OutputKeys.METHOD, "html");
		serializer.transform(domSource, streamResult);
		out.close();

		String result = new String(out.toByteArray());

		BufferedWriter writer = new BufferedWriter(new FileWriter(new File(htmlPath)));
		writer.write(result);
		writer.close();
	}

//	public static void convertPDFToHTML(String pdfPath, String htmlPath) throws IOException, ParserConfigurationException, TransformerException {
//		PDDocument pdf = PDDocument.load(pdfPath);
//		// create the DOM parser
//		PDFDomTree parser = new PDFDomTree();
//		// parse the file
//		parser.processDocument(pdf);
//		// get the DOM Document
//		Document htmlDocument = parser.getDocument();
//
//		ByteArrayOutputStream out = new ByteArrayOutputStream();
//		DOMSource domSource = new DOMSource(htmlDocument);
//		StreamResult streamResult = new StreamResult(out);
//
//		TransformerFactory tf = TransformerFactory.newInstance();
//		Transformer serializer = tf.newTransformer();
//		serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
//		serializer.setOutputProperty(OutputKeys.INDENT, "yes");
//		serializer.setOutputProperty(OutputKeys.METHOD, "html");
//		serializer.transform(domSource, streamResult);
//		out.close();
//
//		String result = new String(out.toByteArray());
//
//		BufferedWriter writer = new BufferedWriter(new FileWriter(new File(htmlPath)));
//		writer.write(result);
//		writer.close();
//	}

}
