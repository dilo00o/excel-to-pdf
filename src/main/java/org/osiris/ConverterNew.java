package org.osiris;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

public class ConverterNew {
    public static String create(File in) {
        StringBuilder builder = new StringBuilder();
        try {
            Workbook workbook = WorkbookFactory.create(in);
            Sheet sheet = workbook.getSheetAt(0);
            HashMap<Integer, List<CellRangeAddress>> cellRangesPerRow = new HashMap<>();
            List<CellRangeAddress> cellRangeAddresses = sheet.getMergedRegions();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            // identify and map all merged regions in the file
            CollectionUtils.emptyIfNull(cellRangeAddresses).forEach(cellrange -> addToMap(cellrange, cellRangesPerRow));
            if (sheet != null) {
                Iterator<Row> rowIt = sheet.rowIterator();
                createHtmlHeader(builder);
                builder.append("<table style=\"border-collapse: collapse;\">");
                while (rowIt.hasNext()) {
                    Row curRow = rowIt.next();
                    List<CellRangeAddress> currMergedCells = cellRangesPerRow.getOrDefault(curRow.getRowNum(), null);
                    if (curRow != null) {
                        Iterator<Cell> cellIterator = curRow.cellIterator();
                        builder.append("<tr>");
                        while (cellIterator.hasNext()) {
                            XSSFCell currCell = (XSSFCell) cellIterator.next();
                            if (isInMergedRegion(currCell, currMergedCells)) {
                                if (currMergedCells.stream().anyMatch(item -> currCell.getColumnIndex() == item.getFirstColumn())) {
                                    CellRangeAddress cc = currMergedCells.stream().filter(item -> currCell.getColumnIndex() == item.getFirstColumn()).findFirst().orElse(new CellRangeAddress(currCell.getRowIndex(), currCell.getRowIndex(), currCell.getColumnIndex(), currCell.getColumnIndex()));
                                    builder.append("<td ")
                                            .append(getCellStyle(currCell))
                                            .append("colspan=\"").append(cc.getLastColumn() - cc.getFirstColumn()).append("\">")
                                            .append(printCell(currCell, evaluator)).append("</td>");
                                }

                            } else { // single cell
                                builder.append("<td ");
                                builder.append(getCellStyle(currCell)).append(">").append(printCell(currCell, evaluator)).append("</td>");
                            }
                        }
                        builder.append("</tr>");
                    }
                }
                builder.append("</table>");
            }
            createHtmlEnd(builder);
            return builder.toString();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return builder.toString();
    }

    private static void createHtmlEnd(StringBuilder builder) {
        builder.append("</body>");
        builder.append("</html>");
        builder.append("</body>");
        builder.append("</html>");
    }

    private static void createHtmlHeader(StringBuilder builder) {
        builder.append("<!DOCTYPE html>");
        builder.append("<html lang=\"en\">");
        builder.append("<head>");
        builder.append("<meta charset=\"UTF-8\">");
        builder.append("<title>Report</title>");
        builder.append("</head>");
        builder.append("<body>");
        builder.append("<!DOCTYPE html>");
        builder.append("<html>");
        builder.append("<head>");
        builder.append("</head>");
        builder.append("<body>");

    }

    private static String getCellStyle(XSSFCell currCell) {
        StringBuilder style = new StringBuilder();
        style.append("style=\" ");
        XSSFCellStyle cellStyle = (XSSFCellStyle) currCell.getCellStyle();
        if (cellStyle != null && cellStyle.getFillForegroundXSSFColor() != null) {
            cellStyle.getFillBackgroundColor();
//            System.out.println(cellStyle.getFillForegroundColor());

            byte[] xfColorByte = cellStyle.getFillForegroundXSSFColor().getRGB();
            if (xfColorByte != null) {
                String coco = String.format("  #%02x%02x%02x;%n", xfColorByte[0], xfColorByte[1], xfColorByte[2]);
                style.append("background-color:").append(coco).append(";");
            }
            byte[] argb = cellStyle.getFillForegroundXSSFColor().getARGB();
            if (argb != null) {
                String coArgb = String.format(" rgba(0x%02x, 0x%02x, 0x%02x, 0x%02x);",
                        argb[3], argb[0], argb[1], argb[2]);
                style.append(coArgb);
            }
            if (cellStyle.getFont() != null) {
                if (cellStyle.getFont().getXSSFColor() != null) {
                    byte[] fontRgb = cellStyle.getFont().getXSSFColor().getRGB();
                    if (fontRgb != null) {
                        String fontColor = String.format("  #%02x%02x%02x;%n", fontRgb[0], fontRgb[1], fontRgb[2]);
                        style.append("color:").append(fontColor).append(";");
                    }
                }
                if (cellStyle.getFont().getBold()) {
                    style.append("font-weight:bold;");
                }
            }
        }
        if (cellStyle != null && cellStyle.getAlignmentEnum() != null) {
            switch (cellStyle.getAlignmentEnum()) {
                case LEFT: {
                    style.append("text-align:left;");
                    break;
                }
                case RIGHT: {
                    style.append("text-align:right;");
                    break;
                }
                case CENTER: {
                    style.append("text-align:center;");
                    break;
                }
                default: {
                    style.append("text-align:center;");
                    break;
                }
            }
        }
        style.append("\" ");
        return style.toString();
    }

    private static boolean isFirstInMergedRegion(XSSFCell currCell, List<CellRangeAddress> currMergedCells) {
        return currMergedCells.stream().anyMatch(item -> currCell.getColumnIndex() == item.getFirstColumn());
    }

    private static boolean isInMergedRegion(XSSFCell currCell, List<CellRangeAddress> currMergedCells) {
        return CollectionUtils.isNotEmpty(currMergedCells) && currMergedCells.stream().anyMatch(item -> item.getFirstColumn() <= currCell.getColumnIndex() && item.getLastColumn() > currCell.getColumnIndex());

    }

    private static String printCell(XSSFCell c, FormulaEvaluator evaluator) {
        String out = "";
        switch (c.getCellTypeEnum()) {
            case ERROR: {
                out = "error";
                break;
            }
            case STRING: {
                out = c.getStringCellValue();
                break;
            }
            case BOOLEAN: {
                out = String.valueOf(c.getBooleanCellValue());
                break;
            }
            case NUMERIC: {
                out = String.valueOf(c.getNumericCellValue());
                break;
            }
            case FORMULA: {
//                out = String.valueOf(c.get)
                Cell cellType = (XSSFCell) evaluator.evaluateInCell(c);
                out = String.valueOf(c.getNumericCellValue());

                break;
            }
            case _NONE: {
                out = "&nbsp;";
                break;
            }
            case BLANK: {
                out = "&nbsp;";
            }

        }
        return out;
    }

    private static void addToMap(CellRangeAddress cellrange, HashMap<Integer, List<CellRangeAddress>> cellRangesPerRow) {
        if (cellRangesPerRow.containsKey(cellrange.getFirstRow())) {
            ArrayList<CellRangeAddress> ll = (ArrayList<CellRangeAddress>) cellRangesPerRow.get(cellrange.getFirstRow());
            ll.add(cellrange);
            cellRangesPerRow.put(cellrange.getFirstRow(), ll);
        } else {
            ArrayList<CellRangeAddress> ll = new ArrayList<>();
            ll.add(cellrange);
            cellRangesPerRow.put(cellrange.getFirstRow(), ll);
        }

    }
}
