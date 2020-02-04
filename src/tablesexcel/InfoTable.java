package tablesexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Formatter;
import java.util.HashMap;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class InfoTable {

	@SuppressWarnings("deprecation")
	public static void main(String args[]) throws IOException {

		// obtaining input bytes from a file
		FileInputStream fis = new FileInputStream(new File(
				"C:/Users/Isen/Desktop/YESS CONSULTING/Dico (1).xls"));

		HSSFWorkbook wb = new HSSFWorkbook(fis);
		// creating a Sheet object to retrieve the object
		HSSFSheet sheet = wb.getSheetAt(0);
		// evaluating cell type
		FormulaEvaluator formulaEvaluator = wb.getCreationHelper()
				.createFormulaEvaluator();
		int noOfRows = sheet.getLastRowNum();
		System.out.println("row size " + noOfRows);
		int noOfColumns = sheet.getRow(0).getLastCellNum();
		ArrayList<Info> infoList = new ArrayList();
		HashMap<String, ArrayList<Info>> infoGroup = new HashMap<String, ArrayList<Info>>();
		String key = "";
		int i = 2;
		while (i < noOfRows) {

			// for (int i = 2; i < noOfRows; i++) {
			System.out.println("iteration " + i + " size item "
					+ infoList.size());
			HSSFRow row = sheet.getRow(i);
			if ((row == null || row.getCell(0) == null || row.getCell(0)
					.getStringCellValue().replace(" ", "").equals(""))) {
				System.out.println("empty ");
				i++;
			} else if (row != null
					&& row.getCell(0).getStringCellValue().contains("<DEBINF")) {
				infoList = new ArrayList<Info>();
				key = row.getCell(0).getStringCellValue()
						.replace("<DEBINF>", "");
				System.out.println("start key " + key);
				i++;

			} else if ((row.getCell(0).getStringCellValue().contains("<FIN"))) {

				infoGroup.put(key, infoList);
				System.out.println("end put " + key);
				i++;
			} else {

				Info info = new Info(
						(getValueOfCell(row.getCell(0)) + "\t\t\t"),
						(getValueOfCell(row.getCell(1)) + "\t\t\t"),
						(getValueOfCell(row.getCell(2)) + "\t\t\t"),
						(getValueOfCell(row.getCell(3)) + "\t\t\t"),
						(getValueOfCell(row.getCell(4)) + "\t\t\t"),
						(getValueOfCell(row.getCell(5)) + "\t\t\t"),
						(getValueOfCell(row.getCell(6))) + "\t\t\t");

				// System.out.println("row.getCell(0) " + row.getCell(0));
				infoList.add(info);
				i++;
			}
		}
		wb.close();
		fis.close();
		System.out.println("end " + infoGroup.size());
		Formatter f1 = new Formatter();
		Formatter f2 = new Formatter();
		Formatter f3 = new Formatter();
		Formatter f4 = new Formatter();
		Formatter f5 = new Formatter();
		Formatter f6 = new Formatter();
		Formatter f7 = new Formatter();
		// left justify
		System.out
				.print("You want to search for an element, if yes enter the key element:");
		Scanner s = new Scanner(System.in);
		String search = s.nextLine();
		while (!search.equals("no") & (!search.equals("exit"))) {
			if (infoGroup.containsKey(search)) {
				System.out
						.printf("|%-36.25s| |%-12.20s| |%-22.20s| |%-22.20s| |%-22.20s| |%-22.20s| |%-22.20s|",
								"label", "code", "format", "longueur",
								" valeur Par Defaut",
								"utilisation Valeur Par Défault",
								"Rubrique Obligatoire");
				System.out.println();
				for (Info info : infoGroup.get(search)) {
					f1 = new Formatter();
					f2 = new Formatter();
					f3 = new Formatter();
					f4 = new Formatter();
					f5 = new Formatter();
					f6 = new Formatter();
					f7 = new Formatter();
					System.out.print(f1.format("%20.25s|", info.label));
					System.out.print(f2.format("%-15.5s", info.code));
					System.out.print(f3.format("|%-7.11s|", info.format));
					System.out.print(f4.format("|%-7.25s|", info.longueur));
					System.out.print(f5.format("%10.25s", info.valParDefaut));
					System.out.print(f6.format("%10.25s", info.utValParDef));
					System.out.print(f7.format("%10.25s", info.rubOblig));

					System.out.println();

				}
				System.out.println();
				System.out
						.println("Continue?Enter another element you want to find or write exit to quit:");
			} else {
				System.out.println("error, again?");
			}
			search = s.nextLine();
		}
	}

	private static String getValueOfCell(HSSFCell cell) {
		if (cell != null) {

			// String key=substringAfter(cell.getStringCellValue(),
			// "<DEBINFO>");

			switch (cell.getCellType()) {
			case HSSFCell.CELL_TYPE_NUMERIC:
				return Double.toString(cell.getNumericCellValue());

			case HSSFCell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
			}

			// infoGroup.put(key, infoList);

		}
		return "";
	}

}


