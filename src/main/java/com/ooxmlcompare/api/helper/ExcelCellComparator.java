package com.ooxmlcompare.api.helper;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import com.ooxmlcompare.api.compare.ComparisonDifference;

public class ExcelCellComparator {

	private List<ComparisonDifference> listOfDifferences;

	public ExcelCellComparator(List<ComparisonDifference> listOfDifferences) {
		this.setListOfDifferences(listOfDifferences);
	}

	public void compareDataInCell(Cell cell1, Cell cell2) {
		if (cell1 == null || cell2 == null) {
			// cell was created or deleted and now POI returns null
			addMessage(cell1.getAddress().toString());
			return;
		}
		if (isCellTypeMatches(cell1, cell2)) {
			final CellType cellType = cell1.getCellType();
			switch (cellType) {
			case BLANK:
			case STRING:
			case ERROR:
				isCellContentMatches(cell1, cell2);
				break;
			case BOOLEAN:
				isCellContentMatchesForBoolean(cell1, cell2);
				break;
			case FORMULA:
				isCellContentMatchesForFormula(cell1, cell2);
				break;
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell1)) {
					isCellContentMatchesForDate(cell1, cell2);
				} else {
					isCellContentMatchesForNumeric(cell1, cell2);
				}
				break;
			default:
				throw new IllegalStateException("Unexpected cell type: " + cellType);
			}
		}

		isCellFillPatternMatches(cell1, cell2);
		isCellAlignmentMatches(cell1, cell2);
		isCellHiddenMatches(cell1, cell2);
		isCellLockedMatches(cell1, cell2);
		isCellFontFamilyMatches(cell1, cell2);
		isCellFontSizeMatches(cell1, cell2);
		isCellFontBoldMatches(cell1, cell2);
		isCellUnderLineMatches(cell1, cell2);
		isCellFontItalicsMatches(cell1, cell2);
		isCellBorderMatches(cell1, cell2, 't');
		isCellBorderMatches(cell1, cell2, 'l');
		isCellBorderMatches(cell1, cell2, 'b');
		isCellBorderMatches(cell1, cell2, 'r');
		isCellFillBackGroundMatches(cell1, cell2);
	}

	/**
	 * Checks if cell alignment matches.
	 */
	private void isCellAlignmentMatches(Cell cell1, Cell cell2) {
		// TODO: check for NPE
		HorizontalAlignment align1 = cell1.getCellStyle().getAlignment();
		HorizontalAlignment align2 = cell2.getCellStyle().getAlignment();
		if (align1 != align2) {
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell border bottom matches.
	 */
	private void isCellBorderMatches(Cell cell1, Cell cell2, char borderSide) {
		if (!(cell1 instanceof XSSFCell))
			return;
		XSSFCellStyle style1 = ((XSSFCell) cell1).getCellStyle();
		XSSFCellStyle style2 = ((XSSFCell) cell2).getCellStyle();
		boolean b1, b2;
		String borderName;
		switch (borderSide) {
		case 't':
		default:
			b1 = style1.getBorderTop() == BorderStyle.THIN;
			b2 = style2.getBorderTop() == BorderStyle.THIN;
			borderName = "TOP";
			break;
		case 'b':
			b1 = style1.getBorderBottom() == BorderStyle.THIN;
			b2 = style2.getBorderBottom() == BorderStyle.THIN;
			borderName = "BOTTOM";
			break;
		case 'l':
			b1 = style1.getBorderLeft() == BorderStyle.THIN;
			b2 = style2.getBorderLeft() == BorderStyle.THIN;
			borderName = "LEFT";
			break;
		case 'r':
			b1 = style1.getBorderRight() == BorderStyle.THIN;
			b2 = style2.getBorderRight() == BorderStyle.THIN;
			borderName = "RIGHT";
			break;
		}
		if (b1 != b2) {
//			addMessage(loc1, loc2, "Cell Border Attributes does not Match ::",
//					(b1 ? "" : "NOT ") + borderName + " BORDER", (b2 ? "" : "NOT ") + borderName + " BORDER");
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell content matches.
	 */
	private void isCellContentMatches(Cell cell1, Cell cell2) {
		// TODO: check for null and non-rich-text cells
		String str1 = cell1.getRichStringCellValue().getString();
		String str2 = cell2.getRichStringCellValue().getString();
		if (!str1.equals(str2)) {
//			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, str1, str2);
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell content matches for boolean.
	 */
	private void isCellContentMatchesForBoolean(Cell cell1, Cell cell2) {
		boolean b1 = cell1.getBooleanCellValue();
		boolean b2 = cell2.getBooleanCellValue();
		if (b1 != b2) {
//			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, Boolean.toString(b1), Boolean.toString(b2));
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell content matches for date.
	 */
	private void isCellContentMatchesForDate(Cell cell1, Cell cell2) {
		Date date1 = cell1.getDateCellValue();
		Date date2 = cell2.getDateCellValue();
		if (!date1.equals(date2)) {
//			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, date1.toGMTString(), date2.toGMTString());
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell content matches for formula.
	 */
	private void isCellContentMatchesForFormula(Cell cell1, Cell cell2) {
		// TODO: actually evaluate the formula / NPE checks
		String form1 = cell1.getCellFormula();
		String form2 = cell2.getCellFormula();
		if (!form1.equals(form2)) {
//			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, form1, form2);
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell content matches for numeric.
	 */
	private void isCellContentMatchesForNumeric(Cell cell1, Cell cell2) {
		// TODO: Check for NaN
		double num1 = cell1.getNumericCellValue();
		double num2 = cell2.getNumericCellValue();
		if (num1 != num2) {
//			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, Double.toString(num1), Double.toString(num2));
			addMessage(cell1.getAddress().toString());
		}
	}

	private String getCellFillBackground(Cell cell) {
		Color col = cell.getCellStyle().getFillForegroundColorColor();
		return (col instanceof XSSFColor) ? ((XSSFColor) col).getARGBHex() : "NO COLOR";
	}

	/**
	 * Checks if cell file back ground matches.
	 */
	private void isCellFillBackGroundMatches(Cell cell1, Cell cell2) {
		String col1 = getCellFillBackground(cell1);
		String col2 = getCellFillBackground(cell2);
		if (!col1.equals(col2)) {
//			addMessage(loc1, loc2, "Cell Fill Color does not Match ::", col1, col2);
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell fill pattern matches.
	 */
	private void isCellFillPatternMatches(Cell cell1, Cell cell2) {
		// TOOO: Check for NPE
		FillPatternType fill1 = cell1.getCellStyle().getFillPattern();
		FillPatternType fill2 = cell2.getCellStyle().getFillPattern();
		if (fill1 != fill2) {
//			addMessage(loc1, loc2, "Cell Fill pattern does not Match ::", Short.toString(fill1), Short.toString(fill2));
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell font bold matches.
	 */
	private void isCellFontBoldMatches(Cell cell1, Cell cell2) {
		if (!(cell1 instanceof XSSFCell))
			return;
		boolean b1 = ((XSSFCell) cell1).getCellStyle().getFont().getBold();
		boolean b2 = ((XSSFCell) cell2).getCellStyle().getFont().getBold();
		if (b1 != b2) {
//			addMessage(loc1, loc2, CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH, (b1 ? "" : "NOT ") + "BOLD",
//					(b2 ? "" : "NOT ") + "BOLD");
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell font family matches.
	 */
	private void isCellFontFamilyMatches(Cell cell1, Cell cell2) {
		// TODO: Check for NPEs
		if (!(cell1 instanceof XSSFCell))
			return;
		String family1 = ((XSSFCell) cell1).getCellStyle().getFont().getFontName();
		String family2 = ((XSSFCell) cell2).getCellStyle().getFont().getFontName();
		if (!family1.equals(family2)) {
//			addMessage(loc1, loc2, "Cell Font Family does not Match ::", family1, family2);
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell font italics matches.
	 */
	private void isCellFontItalicsMatches(Cell cell1, Cell cell2) {
		if (!(cell1 instanceof XSSFCell))
			return;
		boolean b1 = ((XSSFCell) cell1).getCellStyle().getFont().getItalic();
		boolean b2 = ((XSSFCell) cell2).getCellStyle().getFont().getItalic();
		if (b1 != b2) {
//			addMessage(loc1, loc2, CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH, (b1 ? "" : "NOT ") + "ITALICS",
//					(b2 ? "" : "NOT ") + "ITALICS");
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell font size matches.
	 */
	private void isCellFontSizeMatches(Cell cell1, Cell cell2) {
		if (!(cell1 instanceof XSSFCell))
			return;
		short size1 = ((XSSFCell) cell1).getCellStyle().getFont().getFontHeightInPoints();
		short size2 = ((XSSFCell) cell2).getCellStyle().getFont().getFontHeightInPoints();
		if (size1 != size2) {
//			addMessage(loc1, loc2, "Cell Font Size does not Match ::", Short.toString(size1), Short.toString(size2));
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell hidden matches.
	 */
	private void isCellHiddenMatches(Cell cell1, Cell cell2) {
		boolean b1 = cell1.getCellStyle().getHidden();
		boolean b2 = cell1.getCellStyle().getHidden();
		if (b1 != b2) {
//			addMessage(loc1, loc2, "Cell Visibility does not Match ::", (b1 ? "" : "NOT ") + "HIDDEN",
//					(b2 ? "" : "NOT ") + "HIDDEN");
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell locked matches.
	 */
	private void isCellLockedMatches(Cell cell1, Cell cell2) {
		boolean b1 = cell1.getCellStyle().getLocked();
		boolean b2 = cell1.getCellStyle().getLocked();
		if (b1 != b2) {
//			addMessage(loc1, loc2, "Cell Protection does not Match ::", (b1 ? "" : "NOT ") + "LOCKED",
//					(b2 ? "" : "NOT ") + "LOCKED");
			addMessage(cell1.getAddress().toString());
		}
	}

	/**
	 * Checks if cell type matches.
	 */
	private boolean isCellTypeMatches(Cell cell1, Cell cell2) {
		CellType type1 = cell1.getCellType();
		CellType type2 = cell2.getCellType();
		if (type1 == type2)
			return true;
//		addMessage(loc1, loc2, "Cell Data-Type does not Match in :: ", type1.name(), type2.name());
		addMessage(cell1.getAddress().toString());
		return false;
	}

	/**
	 * Checks if cell under line matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell under line matches
	 */
	private void isCellUnderLineMatches(Cell cell1, Cell cell2) {
		// TOOO: distinguish underline type
		if (!(cell1 instanceof XSSFCell))
			return;
		byte b1 = ((XSSFCell) cell1).getCellStyle().getFont().getUnderline();
		byte b2 = ((XSSFCell) cell2).getCellStyle().getFont().getUnderline();
		if (b1 != b2) {
//			addMessage(loc1, loc2, CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH, (b1 == 1 ? "" : "NOT ") + "UNDERLINE",
//					(b2 == 1 ? "" : "NOT ") + "UNDERLINE");
			addMessage(cell1.getAddress().toString());
		}
	}

	private void addMessage(String cellAddress) {
        listOfDifferences.add(new ComparisonDifference(cellAddress, Difference.UPDATE_CELL));
    }
	  
	public List<ComparisonDifference> getListOfDifferences() {
		return listOfDifferences;
	}

	public void setListOfDifferences(List<ComparisonDifference> listOfDifferences) {
		this.listOfDifferences = listOfDifferences;
	}
}
