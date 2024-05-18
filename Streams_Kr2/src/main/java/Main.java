import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {


    public static void main(String[] args) {
        Map<BroadcastsTime, List<Program>> schedule = new TreeMap<>();
        List<Program> allPrograms = new ArrayList<>();

        try (BufferedReader br = Files.newBufferedReader(Paths.get("data.txt"))) {
            String channel = "";
            BroadcastsTime time = null;
            String title = "";
            String line;

            while ((line = br.readLine()) != null) {
                line = line.trim();
                if (line.startsWith("#")) {

                } else if (line.isEmpty()) {

                } else {
                    if (channel.isEmpty()) {
                        channel = line;
                    } else if (time == null) {
                        time = new BroadcastsTime(line);
                    } else {
                        title = line;
                        Program program = new Program(channel, time, title);
                        allPrograms.add(program);
                        schedule.computeIfAbsent(time, k -> new ArrayList<>()).add(program);
                        channel = "";
                        time = null;
                        title = "";
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }


        allPrograms.sort(Comparator.comparing(Program::getChannel)
                .thenComparing(Program::getTime));

        for (Program program : allPrograms) {
            System.out.println(program);
        }

        BroadcastsTime now = new BroadcastsTime(LocalTime.now().format(DateTimeFormatter.ofPattern("HH:mm")));
        for (Program program : allPrograms) {
            if (program.getTime().equals(now)) {
                System.out.println("Сейчас идут: " + program);
            }
        }


        String searchTitle = "Все на Матч!";
        for (Program program : allPrograms) {
            if (program.getTitle().equalsIgnoreCase(searchTitle)) {
                System.out.println("Поиск программы: " + program);
            }
        }

        String searchChannel = "Первый";
        for (Program program : allPrograms) {
            if (program.getChannel().equalsIgnoreCase(searchChannel) && program.getTime().equals(now)) {
                System.out.println("Идут сейчас на канале " + searchChannel + ": " + program);
            }
        }

        BroadcastsTime startTime = new BroadcastsTime("04:45");
        BroadcastsTime endTime = new BroadcastsTime("05:20");
        for (Program program : allPrograms) {
            if (program.getChannel().equalsIgnoreCase(searchChannel) &&
                    program.getTime().between(startTime, endTime)) {
                System.out.println("Программа на " + searchChannel + " в " + startTime + " - " + endTime + ": " + program);
            }
        }

        saveToExcel(allPrograms);
    }

    private static void saveToExcel(List<Program> programs) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Programs");

        int rowCount = 0;
        for (Program program : programs) {
            Row row = sheet.createRow(rowCount++);
            row.createCell(0).setCellValue(program.getChannel());
            row.createCell(1).setCellValue(program.getTime().toString());
            row.createCell(2).setCellValue(program.getTitle());
        }

        try (FileOutputStream outputStream = new FileOutputStream("programs.xlsx")) {
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }



}
