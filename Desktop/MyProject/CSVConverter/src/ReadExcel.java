import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import org.apache.commons.lang.StringUtils;
//import au.com.bytecode.opencsv.CSVReader;


import java.sql.Statement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.StringTokenizer;
import java.sql.ResultSet;

public class ReadExcel {
	static String value = "";
	static String Final = "";
	static int rowNum = 0;
	static int colNum = 0;
	static int count = 0;
	static String csvFile = "/home/shinto/Desktop/shinto11.csv";
	static String columnvaluesonebyone = "";

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Parser();// For reading the excel file
		generateCsvFile(csvFile);// Generate CSV file and insert 'final' into
									// it.
		ReadFromCSV();// To display the CSV in a format
		SaveCSVtoDB();// For Deciding what to do with the CSV - Insert to
							// DB.
	}

	// For Reading the file
	public static String Parser() throws IOException {
		File excel = new File("/home/shinto/Desktop/shinto.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet ws = wb.getSheet("Sheet1");
		rowNum = ws.getLastRowNum() + 1;
		colNum = ws.getRow(0).getLastCellNum();
		String[][] data = new String[rowNum][colNum];
		for (int i = 0; i < rowNum; i++) {
			XSSFRow row = ws.getRow(i);
			for (int j = 0; j < colNum; j++) {
				XSSFCell cell = row.getCell(j);
				value = cell.toString();
				data[i][j] = value;
				Final = Final.concat(value + ",");
			}
		}
		return (Final);
	}

	// Create a CSV and move the data to it
	public static void generateCsvFile(String sFileName) throws IOException {
		try {
			FileWriter writer = new FileWriter(sFileName);
			writer.append(Final);
			writer.flush();
			writer.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	// To extract data in correct format from CSV
	public static void ReadFromCSV() throws Exception {
		Connection con = null;
		BufferedReader br = new BufferedReader(new FileReader(csvFile));
		String line = "";
		StringTokenizer st = null;
		int lineNumber = 0, tokenNumber = 0;

		while ((line = br.readLine()) != null) {
			lineNumber++;
			// use comma as token separator
			st = new StringTokenizer(line, ",");
			while (st.hasMoreTokens()) {
				tokenNumber++;
				count++;				
				if (count > colNum) {
					System.out.println();
					System.out.println("=============================");
					count = 0;
				} else {
					columnvaluesonebyone = st.nextToken();
					System.out.print(columnvaluesonebyone + "   ");
				}
			}
		}
		System.out.println();	
		tokenNumber = 0;
	}

	// For manipulating the CSV Data.
	public static void SaveCSVtoDB() throws Exception {
		Class.forName("com.mysql.jdbc.Driver");
		Connection con = null;
		PreparedStatement stmt = null;
		BufferedReader br = new BufferedReader(new FileReader(csvFile));
		String line = br.readLine();
		String[] rowItems = line.split(",");
		try {
			con = DriverManager.getConnection(
					"jdbc:mysql://localhost:3306/NewDB", "root", "qburst");
			stmt = con
					.prepareStatement("insert into Student (Name,Mark1,Mark2,Mark3)values(?,?,?,?)");
			int setstringnumber = 0;
			for (int i = colNum; i < rowItems.length; i++) {
				if (i % colNum == 0) {
					stmt.setString(setstringnumber + 1, rowItems[i]);
					setstringnumber++;
				} else {
					stmt.setFloat(setstringnumber + 1,
							Float.parseFloat(rowItems[i]));
					setstringnumber++;
				}
				if (setstringnumber == colNum) {
					stmt.executeUpdate();
					setstringnumber = 0;
				}
			}
			System.out.println("=============================");
			System.out.println("Successfully Inserted!!");
			System.out.println("=============================");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (stmt != null) {
				try {
					stmt.close();
				} catch (SQLException ex) {
				}
			}
			if (con != null) {
				try {
					con.close();
				} catch (SQLException ex) {
				}
			}
		}

	}
}
