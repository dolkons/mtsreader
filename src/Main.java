import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.nio.charset.Charset;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by Константин on 11.09.2016.
 */
public class Main {
	private static final String out = "out";
	private static final long startDateUnixTime = 315964800; //6 января 1980 года в Unixtime
	private static final Date startDate = new Date(startDateUnixTime);
	private static final long secondsInWeek = 604800; //число секунд в неделе
	private static final long secondsInHour = 3600;
	private static final DateFormat df = new SimpleDateFormat("dd/MM/yyyy HH:mm");
	private static final int[] means = {5, 10};
	private static int cellCount;

	private static int magneticX;
	private static int magneticY;
	private static int magneticZ;

	private static int meanTxt5Min;
	private static int meanTxt10Min;

	private static String mainDir;

	public static void main(String[] args) throws IOException, ParseException {

		Properties prop = new Properties();
		InputStream input = null;

		input = new FileInputStream("./config.properties");
		//input = new FileInputStream(System.getProperty("user.dir")+"/config.properties");
		prop.load(input);

		magneticX = Integer.parseInt(prop.getProperty("magneticX"));
		magneticY = Integer.parseInt(prop.getProperty("magneticY"));
		magneticZ = Integer.parseInt(prop.getProperty("magneticZ"));

		meanTxt5Min = Integer.parseInt(prop.getProperty("meanTxt5Min"));
		meanTxt10Min = Integer.parseInt(prop.getProperty("meanTxt10Min"));

		mainDir = prop.getProperty("DataDir");

		df.setTimeZone(TimeZone.getTimeZone("GMT"));
		File folder = new File(mainDir);
		File[] listOfFiles = folder.listFiles();

		for (int i = 0; i < listOfFiles.length; i++) {
			if (listOfFiles[i].isFile()) {
				System.out.println("File: " + listOfFiles[i].getName());
			} else if (listOfFiles[i].isDirectory()) {
				System.out.println("Directory: " + listOfFiles[i].getName());
				if (listOfFiles[i].getName().startsWith("MTS")) {
					File mtsDataFolder = new File(mainDir + listOfFiles[i].getName());
					File[] dataFilesInMtsFolder = mtsDataFolder.listFiles();
					File outDirectory = new File(mainDir + out);
					if (!outDirectory.exists()) {
						outDirectory.mkdir();//create "out" directory
					}
					writeToExcel(mtsDataFolder.getName(), dataFilesInMtsFolder, mainDir + out + "/xls/" + mtsDataFolder.getName() + ".xls");
				}
			}
		}
	}

	private static void writeToExcel(String folderName, File[] dataFiles, String outFileName) throws IOException, ParseException {

		int rowPosition = 1;
		File outFileDirectory = new File(mainDir + out + "/xls");
		if (!outFileDirectory.exists()) {
			outFileDirectory.mkdir();
		}

		Workbook book = new HSSFWorkbook();
		Sheet sheet = book.createSheet(folderName);
		Row row;
		Row header = sheet.createRow(0);
		cellCount = 0;
		for (int meanFactor : means) {
			Cell date = header.createCell(cellCount);
			date.setCellValue("Дата");

			Cell magneticXTitle = header.createCell(1 + cellCount);
			magneticXTitle.setCellValue("magneticX");

			Cell magneticYTitle = header.createCell(2 + cellCount);
			magneticYTitle.setCellValue("magneticY");

			Cell magneticZTitle = header.createCell(3 + cellCount);
			magneticZTitle.setCellValue("magneticZ");

			for (File dataFile : dataFiles) {
				ArrayList<String> lines = new ArrayList<>();
				InputStream fis = new FileInputStream(dataFile.getAbsolutePath());
				InputStreamReader isr = new InputStreamReader(fis, Charset.forName("UTF-8"));
				BufferedReader br = new BufferedReader(isr);
				String line;
				while ((line = br.readLine()) != null) {//цикл по каждой строке в файле
					lines.add(line);
				}
				br.close();
				String dateString = getDateFromWeekCount(folderName, dataFile.getName());
				switch (meanFactor) {
					case 5: {
						row = sheet.createRow(rowPosition);
						row.createCell(cellCount).setCellValue(dateString);
						rowPosition = writeMean(sheet, dateString, lines, rowPosition, meanTxt5Min);//усреднение по 5 минутам. (частота дискр - 1 гц)
						break;
					}
					case 10: {
						row = sheet.getRow(rowPosition);
						row.createCell(cellCount).setCellValue(dateString);
						rowPosition = writeMean(sheet, dateString, lines, rowPosition, meanTxt10Min);//усреднение по 10 минутам.
						break;
					}
				}
			}
			rowPosition = 1;
			sheet.autoSizeColumn(cellCount);
			cellCount += 14;
		}
		File outFile = new File(outFileName);
		System.out.println(outFileName);

		book.write(new FileOutputStream(outFile));
		book.close();
	}

	private static int writeMean(Sheet sheet, String dateString, ArrayList<String> lines, int rowPosition, int meanFactor) throws ParseException {

		double magneticXSum = 0;
		double magneticYSum = 0;
		double magneticZSum = 0;

		int lineCount = 1;//счетчик строк в исходном файле с данными
		int validStringsCount = 0; //счетчик валидных строк (исп для усреднения только по корректным строкам)
		int totalLineCount = rowPosition;//счетчик строк в ексель файле
		int meanCount = 0;//счетчик усреднений
		Row row;
		for (String line : lines) {
			line = line.trim();
			String[] values = line.split("\\s+");

			if (!isValidString(values)){ //проверка на нулевую строку или строку, содержащую слишком высокие значения
				continue;
			}
			else {
				magneticXSum += (Double.parseDouble(values[0]) / magneticX);
				magneticYSum += (Double.parseDouble(values[1]) / magneticY);
				magneticZSum += (Double.parseDouble(values[2]) / magneticZ);

				validStringsCount++;
			}

			if (lineCount == meanFactor) {
				row = sheet.getRow(totalLineCount);
				if (row == null) {
					row = sheet.createRow(totalLineCount);
				}

				Date date = df.parse(dateString);
				Calendar calendar = Calendar.getInstance();
				calendar.setTimeInMillis(date.getTime() + meanCount * getMinuteFromMeanFactor(meanFactor) * 60 * 1000);
				row.createCell(cellCount).setCellValue(df.format(calendar.getTime()));

				if (validStringsCount == 0){
					Cell magneticXCell = row.createCell(1 + cellCount);
					magneticXCell.setCellValue(0);
					Cell magneticYCell = row.createCell(2 + cellCount);
					magneticYCell.setCellValue(0);
					Cell magneticZCell = row.createCell(3 + cellCount);
					magneticZCell.setCellValue(0);
				}
				else {
					Cell magneticXCell = row.createCell(1 + cellCount);
					magneticXCell.setCellValue((magneticXSum + Double.parseDouble(values[0]) / magneticX) / validStringsCount);
					Cell magneticYCell = row.createCell(2 + cellCount);
					magneticYCell.setCellValue((magneticYSum + Double.parseDouble(values[1]) / magneticY) / validStringsCount);
					Cell magneticZCell = row.createCell(3 + cellCount);
					magneticZCell.setCellValue((magneticZSum + Double.parseDouble(values[2]) / magneticZ) / validStringsCount);
				}

				magneticXSum = 0;
				magneticYSum = 0;
				magneticZSum = 0;

				lineCount = 0;
				validStringsCount = 0;
				totalLineCount++;
				meanCount++;
			}
			lineCount++;
		}
		return totalLineCount;
	}

	private static boolean isValidString(String[] values) {
		if (values.length < 3) {
			return false;
		}
		boolean allZeros = true;
		boolean brokenData = false;
		for (String value : values){
			try {
				if (Double.parseDouble(value) != 0 && Double.parseDouble(value) < Math.pow(2, 23)){
					allZeros = false;
				}
				else{
					allZeros = true;
				}
			}
			catch (NumberFormatException e) {
				brokenData = true;
			}
		}
		return (!allZeros&&!brokenData);
	}

	private static String getDateFromWeekCount(String folderName, String fileName) {
		int weekCount = Integer.parseInt(folderName.split("\\.")[0].split("MTS")[1]);
		int hourInWeek = Integer.parseInt(fileName.split("\\.")[0].split("HOUR")[1]);
		long weekInUnixTime = weekCount * secondsInWeek + hourInWeek * secondsInHour + startDateUnixTime;
		Calendar calendar = Calendar.getInstance();
		calendar.setTimeInMillis(weekInUnixTime * 1000);
		return df.format(calendar.getTime());
	}

	private static int getMinuteFromMeanFactor(int meanFactor) {
		int minute;
		switch (meanFactor) {
			case 15000: {
				minute = 5;
				break;
			}
			case 30000: {
				minute = 10;
				break;
			}
			case 300: {
				minute = 5;
				break;
			}
			case 600:{
				minute = 10;
				break;
			}
			default: {
				minute = 0;
				break;
			}
		}
		return minute;
	}
}
