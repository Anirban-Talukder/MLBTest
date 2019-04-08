package com.test.mlb;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import jxl.CellType;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class SecretSantaGame {

	static ArrayList<String> players = new ArrayList<String>();
	static ArrayList<String> giftee;
	static ArrayList<String> santa = new ArrayList<String>();
	static File file = null;
	static WritableWorkbook myReport = null;
	static WritableSheet mySheet = null;
	static Label[] label = null;

	// Adding players to the game
	public static ArrayList<String> addPlayers(String... names) {
		for (String name : names) {
			if (players.contains(name)) {
				System.out.println("Couldn't add " + name + " again, as already registered");
				// break;
			} else
				players.add(name);
		}
		return players;
	}

	// Storing giftees name to data file
	public static void storeData(ArrayList recv, int m) throws Exception {
		Label[] label = new Label[players.size()];
		File file = new File("C:\\Users\\aurph\\eclipse-workspace\\MLBTest\\Utility\\excell.xls");
		WritableWorkbook myReport = Workbook.createWorkbook(file);
		WritableSheet mySheet = myReport.createSheet("MySheet", 0);
		for (int p = 0; p < m; p++) {
			for (int i = 0; i < recv.size(); i++) {
				label[i] = new Label(p, i, (String) recv.get(i));
				mySheet.addCell(label[i]);
			}
		}
		myReport.write();
		myReport.close();
	}

	public static void readData(int i, int j) throws Exception {
		Workbook wb = Workbook
				.getWorkbook(new File("C:\\Users\\aurph\\eclipse-workspace\\MLBTest\\Utility\\excell.xls"));
		String data00 = wb.getSheet(0).getCell(i, j).getContents();
		System.out.println(data00);
	}

	public static void modifyData() throws Exception {
		giftee = (ArrayList<String>) players.clone();
		Workbook wb = Workbook
				.getWorkbook(new File("C:\\Users\\aurph\\eclipse-workspace\\MLBTest\\Utility\\excell.xls"));
		WritableWorkbook copy = Workbook
				.createWorkbook(new File("C:\\Users\\aurph\\eclipse-workspace\\MLBTest\\Utility\\excell.xls"), wb);
		mySheet = copy.getSheet(0);

		Label[] label = new Label[players.size()];
		WritableCell cell = null;
		for (int i = 0; i < players.size(); i++) {
			for (int j = 0; j < players.size(); j++) {
				cell = mySheet.getWritableCell(j, i);
				if (cell.getType() == CellType.LABEL) {

					label[i] = (Label) cell;
					label[i].setString(players.get(i));
				}
			}
		}
		copy.write();
		copy.close();
	}

	// Playing the game
	public static void play() throws Exception {

		giftee = (ArrayList<String>) players.clone();
		// Shuffling actual list
		Collections.shuffle(giftee);
		for (int i = 0; i < players.size(); i++) {
			int target = 0;
			// Checking if receiver and giftee are same
			if (giftee.get(target).equals(players.get(i))) {
				if (giftee.size() > 1) {
					target++;
				}
			}

			System.out.println(players.get(i) + "'s Secret Santa is  => " + giftee.get(target));
			// Storing receivers name to data before removal
			santa.add(giftee.get(target));
			// removing receiver to erase duplication
			giftee.remove(giftee.get(target));

		}
		// storeData(santa, 1);
	}

	public static void main(String[] args) throws Exception {
		addPlayers("Moni", "Smriti", "Ayon", "Anom", "Ahon", "Hashu", "Mishu", "Ayon");
//		System.out.println(players);
		play();
		storeData(santa, 3);
		readData(1, 1);
		// modifyData();

	}
}