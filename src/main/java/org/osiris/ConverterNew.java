package org.osiris;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
                                    CellRangeAddress cc = currMergedCells.stream().filter(item -> currCell.getColumnIndex() == item.getFirstColumn()).findFirst().orElse(new CellRangeAddress(currCell.getRowIndex(), currCell.getRowIndex(), currCell.getColumnIndex(), currCell.getColumnIndex()));
                                    builder.append("<td colspan=\"").append(cc.getLastColumn() - cc.getFirstColumn()).append("\">").append(printCell(currCell, evaluator)).append("</td>");
                                }

                            } else { // single cell
                                builder.append("<td> ").append(printCell(currCell, evaluator)).append("</td>");
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
