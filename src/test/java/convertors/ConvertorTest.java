package convertors;

import java.io.FileNotFoundException;
import java.io.IOException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import org.apache.pdfbox.exceptions.CryptographyException;

import junit.framework.TestCase;

public class ConvertorTest extends TestCase {
	public void testDoc() throws FileNotFoundException, ParserConfigurationException, IOException, TransformerException {
		String docFilePath = "https://192.168.2.100:8443/contentserver/warehouse/3/resumes/ananthan.doc";
		long st = System.currentTimeMillis();
		String html = Convertor.convertDocToHTML(docFilePath);
		long et = System.currentTimeMillis();
		System.out.println("convertDocToHTML Time taken: " + (et - st));
		assertNotNull(html);
	}
	public void testDoc2() throws FileNotFoundException, ParserConfigurationException, IOException, TransformerException {
		String docFilePath = "C:\\Users\\Admin\\Downloads\\Sample Resumes\\doc\\CVSample.doc";
		String html = Convertor.convertDocToHTML(docFilePath);
		assertNotNull(html);
	}
	
	

	public void testDocx() throws FileNotFoundException, IOException {
		String docxFilePath = "C:\\Users\\Admin\\Downloads\\Sample Resumes\\docx\\MBA Marketing Fresher Resume Sample Doc.docx";
		long st = System.currentTimeMillis();
		String html = Convertor.convertDocxToHTML(docxFilePath);
		long et = System.currentTimeMillis();
		System.out.println("convertDocxToHTML Time taken: " + (et - st));
		assertNotNull(html);
	}

	public void testPdf() throws IOException, ParserConfigurationException, TransformerException, CryptographyException {
		String pdfFilePath = "C:\\Users\\Admin\\Downloads\\Sample Resumes\\pdf\\samples.pdf";
		long st = System.currentTimeMillis();
		String html = Convertor.convertPDFToHTML(pdfFilePath);
		long et = System.currentTimeMillis();
		System.out.println("convertPDFToHTML Time taken: " + (et - st));
		assertNotNull(html);
	}
	

}
