package convertors;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.io.FilenameUtils;
import org.apache.pdfbox.exceptions.CryptographyException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.fit.pdfdom.PDFDomTree;
import org.w3c.dom.Document;

import convertors.utility.HttpsHandler;

/**
 * @author Kurra Raghu
 *
 */
public final class Convertor {
	private Convertor() {

	}

	public static String getHtml(String fileUrlorPath) throws Exception {
		String exten = FilenameUtils.getExtension(fileUrlorPath);
		String html = null;
		try {
			switch (exten) {
			case "docx":
				html = Convertor.convertDocxToHTML(fileUrlorPath);
				break;
			case "doc":
				html = convertDocToHTML(fileUrlorPath);
				break;
			case "pdf":
				html = Convertor.convertPDFToHTML(fileUrlorPath);
				break;
			default:
				throw new FileNotFoundException("file not sipported for parsing: " + fileUrlorPath);
			}
		} catch (Exception e) {
			throw e;
		}
		return html;
	}

	/**
	 * @author Kurra Raghu
	 * @param docxFilePath
	 * @return HTMLString
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static String convertDocxToHTML(String docxFilePath) throws FileNotFoundException, IOException {
		ByteArrayOutputStream out = null;
		String docxHtml = null;
		try {
			XWPFDocument document = new XWPFDocument(getFileInputStream(docxFilePath));
			XHTMLOptions options = XHTMLOptions.create();// .indent( 4 );
			// // Extract image
			// File imageFolder = new File(root + "/images/" + fileInName);
			// options.setExtractor(new FileImageExtractor(imageFolder));
			// // URI resolver
			// options.URIResolver(new FileURIResolver(imageFolder));
			out = new ByteArrayOutputStream();
			XHTMLConverter.getInstance().convert(document, out, options);
			docxHtml = new String(out.toByteArray());
			FileOutputStream fos = new FileOutputStream(new File("E:\\myFile.doc"));
			out.writeTo(fos);
		} finally {
			if (out != null) {
				out.close();
			}
		}
		return docxHtml;
	}

	/**
	 * @author Kurra Raghu
	 * @param docFilePath
	 * @return HTMLString
	 * @throws ParserConfigurationException
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws TransformerException
	 */
	public static String convertDocToHTML(String docFilePath) throws ParserConfigurationException, FileNotFoundException, IOException, TransformerException {
		ByteArrayOutputStream out = null;
		String docHtml = null;
		try {
			HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(getFileInputStream(docFilePath));
			WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
			wordToHtmlConverter.processDocument(wordDocument);
			Document htmlDocument = wordToHtmlConverter.getDocument();
			DOMSource domSource = new DOMSource(htmlDocument);
			out = new ByteArrayOutputStream();
			StreamResult streamResult = new StreamResult(out);
			TransformerFactory tf = TransformerFactory.newInstance();
			Transformer serializer = tf.newTransformer();
			serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");
			serializer.setOutputProperty(OutputKeys.METHOD, "html");
			serializer.transform(domSource, streamResult);
			docHtml = new String(out.toByteArray());
		} finally {
			if (out != null) {
				out.close();
			}
		}
		return docHtml;
	}

	/**
	 * @author Kurra Raghu
	 * @param pdfFilePath
	 * @return HTMLString
	 * @throws IOException
	 * @throws ParserConfigurationException
	 * @throws TransformerException
	 * @throws CryptographyException
	 */
	public static String convertPDFToHTML(String pdfFilePath) throws IOException, ParserConfigurationException, TransformerException, CryptographyException {
		ByteArrayOutputStream out = null;
		String pdfHtml = null;
		try {
			PDDocument pdf = PDDocument.load(getFileInputStream(pdfFilePath));
			if (pdf.isEncrypted()) {
				pdf.decrypt("");
			}
			// create the DOM parser
			PDFDomTree parser = new PDFDomTree();
			parser.setDisableImageData(true);
			// parse the file
			parser.processDocument(pdf);
			// get the DOM Document
			Document htmlDocument = parser.getDocument();

			DOMSource domSource = new DOMSource(htmlDocument);
			out = new ByteArrayOutputStream();
			StreamResult streamResult = new StreamResult(out);
			TransformerFactory tf = TransformerFactory.newInstance();
			Transformer serializer = tf.newTransformer();
			serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");
			serializer.setOutputProperty(OutputKeys.METHOD, "html");
			serializer.transform(domSource, streamResult);
			pdfHtml = new String(out.toByteArray());
		} finally {
			if (out != null) {
				out.close();
			}
		}
		return pdfHtml;
	}

	/**
	 * @author Kurra Raghu
	 * @param fileUrlorPath
	 * @return InputStream
	 * @throws FileNotFoundException
	 */
	private static InputStream getFileInputStream(String fileUrlorPath) throws FileNotFoundException {
		try {
			URL url = new URL(fileUrlorPath);
			return url.openConnection().getInputStream();
		} catch (Exception e) {
			try {
				//Disable HTTPS and get Response
				URL url = new URL(fileUrlorPath);
				return HttpsHandler.getInputStream(url);
			} catch (Exception e1) {
				//If fileUrlorPath is Not a usl path and is a file location path
				return new FileInputStream(fileUrlorPath);
			}
		}
	}

}
