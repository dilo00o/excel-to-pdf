package org.osiris;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetUtil;
import org.apache.poi.xssf.usermodel.*;
import org.xhtmlrenderer.css.style.derived.StringValue;

import java.io.File;
import java.io.IOException;
import java.util.*;

public class ConverterNew {
    public static void create(File in, String out) {
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
                builder.append("<table>");
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
//                                    System.out.println("cell: " + printCell(currCell, evaluator) + "  backgroundcolor: " + currCell.getCellStyle().getFillBackgroundXSSFColor());
//                                    System.out.println("cell: " + printCell(currCell, evaluator) + "  font: " + currCell.getCellStyle().getFont().getXSSFColor());
                                    CellRangeAddress cc = currMergedCells.stream().filter(item -> currCell.getColumnIndex() == item.getFirstColumn()).findFirst().orElse(new CellRangeAddress(currCell.getRowIndex(), currCell.getRowIndex(), currCell.getColumnIndex(), currCell.getColumnIndex()));
                                    builder.append("<td ")
                                            .append(getCellStyle(currCell))
                                            .append("colspan=\"").append(cc.getLastColumn() - cc.getFirstColumn()).append("\">")
                                            .append(printCell(currCell, evaluator)).append("</td>");
                                }

                            } else { // single cell
//                                System.out.println("cell: " + printCell(currCell, evaluator) + "  backgroundcolor: " + currCell.getCellStyle().getFillBackgroundColor());
//                                System.out.println("cell: " + printCell(currCell, evaluator) + "  foreground: " + currCell.getCellStyle().getFillForegroundColor());
                                builder.append("<td ");
                                builder.append(getCellStyle(currCell)).append(">").append(printCell(currCell, evaluator)).append("</td>");
                            }
                        }
                        builder.append("</tr>");
                    }
                }
                builder.append("</table>");
            }
            String ooo = builder.toString();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private static String getCellStyle(XSSFCell currCell) {
        StringBuilder style = new StringBuilder();
        style.append("style=\"");
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

    private class LocalCell {
        private String value;
        private int colSpan;
        private String color;
        private String fontBold;

        public LocalCell() {
        }

        public String getValue() {
            return value;
        }

        public void setValue(String value) {
            this.value = value;
        }

        public int getColSpan() {
            return colSpan;
        }

        public void setColSpan(int colSpan) {
            this.colSpan = colSpan;
        }

        public String getColor() {
            return color;
        }

        public void setColor(String color) {
            this.color = color;
        }

        public String getFontBold() {
            return fontBold;
        }

        public void setFontBold(String fontBold) {
            this.fontBold = fontBold;
        }
    }
}
