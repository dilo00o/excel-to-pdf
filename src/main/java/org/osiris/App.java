package org.osiris;

import com.itextpdf.forms.PdfPageFormCopier;
import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.html2pdf.css.media.MediaDeviceDescription;
import com.itextpdf.html2pdf.css.media.MediaType;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;

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
//                obj.getFile(args[1], args[2]);
                obj.getFileNew(args[1], args[2]);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            try {
//                obj.getFile("xl.xlsx");
                obj.getFileNew("xl.xlsx", "putzo");
//				obj.mergePdf("sample-merge.pdf");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private void getFile(String fileName) throws IOException {
        getFile(fileName, "sample_out");
    }

    private void getFileNew(String fileName, String outFileName) throws IOException {
        URL resource = getClass().getClassLoader().getResource(fileName);
        File in = null;
        if (resource != null) {
            in = new File(resource.getFile());
            ConverterNew.create(in, new String());
        }
    }

    private void getFile(String fileName, String outFileName) throws IOException {
        URL resource = getClass().getClassLoader().getResource(fileName);
        File in = null;
        if (resource != null) {
            in = new File(resource.getFile());
            ConverterNew.create(in, new String());
            StringWriter printWriter = new StringWriter();
            ExcelToHtmlConverter toHtml = ExcelToHtmlConverter.create(in, printWriter);
            toHtml.setCompleteHTML(true);
            toHtml.generateHtml();
            String outHtml = printWriter.toString();
            // get rid of euro sign and encode it to html
            outHtml = outHtml.replaceAll("â‚¬", "&#8364;").replaceAll("€", "&#8364;");
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            PageSize pageSize = PageSize.A4.rotate();
            PdfDocument pdfWriter = new PdfDocument(new PdfWriter(out));
            pdfWriter.setDefaultPageSize(pageSize);
            ConverterProperties properties = new ConverterProperties();
            properties.setCharset("UTF-8");
            MediaDeviceDescription mediaDeviceDescription
                    = new MediaDeviceDescription(MediaType.SCREEN);
            mediaDeviceDescription.setWidth(pageSize.getWidth());
            mediaDeviceDescription.setOrientation("landscape");
            properties.setMediaDeviceDescription(mediaDeviceDescription);
//			HtmlConverter.convertToPdf(outHtml, out, properties);
            HtmlConverter.convertToPdf(outHtml, pdfWriter, properties);

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

    private void mergePdf(String dest) throws IOException {
        PdfDocument pdfDoc = new PdfDocument(new PdfReader(getClass().getClassLoader().getResource("nnn.pdf").getFile()), new PdfWriter(dest));
        PdfDocument cover = new PdfDocument(new PdfReader(getClass().getClassLoader().getResource("bazinga.pdf").getFile()));
        cover.copyPagesTo(1, 1, pdfDoc, 1, new PdfPageFormCopier());
        cover.close();
        pdfDoc.close();
    }
}
