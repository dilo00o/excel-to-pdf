package org.osiris;

import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.kernel.geom.PageSize;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

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

		try {
//			obj.getFile("budget-breakdown.xlsx");
			obj.getFile("report.xlsx");
		} catch (IOException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
	}

	private void getFile(String fileName) throws IOException, InvalidFormatException {

		// Creating a Workbook from an Excel file (.xls or .xlsx)
		URL resource = getClass().getClassLoader().getResource(fileName);
		File in = new File(resource.getFile());
//		PrintWriter printWriter = new PrintWriter(new FileWriter("./bazinga.html"));
		StringWriter printWriter = new StringWriter();
		ExcelToHtmlConverter toHtml = ExcelToHtmlConverter.create(in, printWriter);
		toHtml.setCompleteHTML(true);
		toHtml.generateHtml();
		File pdfDest = new File("./bazinga.pdf");

		OutputStream os = new FileOutputStream(pdfDest);
		HtmlConverter.convertToPdf(printWriter.toString(), os);

	}
}
