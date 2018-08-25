package org.osiris;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Formatter;

public class UTIL {
	public static void setupColorMap(Workbook wb, HtmlHelper helper) {
		if (wb instanceof HSSFWorkbook) {
			helper = new HSSFHtmlHelper((HSSFWorkbook) wb);
		} else if (wb instanceof XSSFWorkbook) {
			helper = new XSSFHtmlHelper();
		} else {
			throw new IllegalArgumentException(
					"unknown workbook type: " + wb.getClass().getSimpleName());
		}
	}

	public static void ensureOut(Formatter out, Appendable output) {
		if (out == null) {
			out = new Formatter(output);
		}
	}

}
