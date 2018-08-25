package org.osiris;

import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.kernel.geom.PageSize;

import java.io.*;
import java.net.URL;

/**
 * Hello world!
 */
public class App {
	public static final PageSize[] pageSizes = {
			PageSize.A4.rotate(),
			new PageSize(720, PageSize.A4.getHeight()),
			new PageSize(PageSize.A5.getWidth(), PageSize.A4.getHeight())
	};

	public static void main(String[] args) {
		App obj = new App();
		if (args.length == 3) {
			try {
				obj.getFile(args[1], args[2]);
			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			try {
				obj.getFile("xl.xlsx");
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	private void getFile(String fileName) throws IOException {
		getFile(fileName, "bazinga");
	}

	private void getFile(String fileName, String outFileName) throws IOException {
		URL resource = getClass().getClassLoader().getResource(fileName);
		File in = new File(resource.getFile());
		StringWriter printWriter = new StringWriter();
		ExcelToHtmlConverter toHtml = ExcelToHtmlConverter.create(in, printWriter);
		toHtml.setCompleteHTML(true);
		toHtml.generateHtml();
		String outHtml = printWriter.toString();
		// get rid of euro sign and encode it to html
		outHtml = outHtml.replaceAll("â‚¬", "&#8364;").replaceAll("€", "&#8364;");
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		HtmlConverter.convertToPdf(outHtml, out);
		FileOutputStream fos = null;
		try { // just for testing
			fos = new FileOutputStream(new File("./" + outFileName + ".pdf"));
			out.writeTo(fos);
		} catch (IOException ioe) {
			// Handle exception here
			ioe.printStackTrace();
		} finally {
			fos.close();
		}
//		DataHandler dataHandler = new DataHandler(out.toByteArray(), "application/pdf");
// return this

	}
}
