
import java.text.SimpleDateFormat;
import java.io.*;
import java.util.*;
import de.re.easymodbus.server.*;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class modbus {
    static int period=1000;
    public static void main(String[] args) throws IOException {

        Timer timer = new Timer();
        TimerTask task = new MyTimerTask(timer);

        timer.schedule(task, 0, period);
        System.out.println("TimerTask started");

    }
}

class MyTimerTask extends TimerTask {
    static int num = 0;
    
    static HSSFWorkbook workbook;
    static HSSFSheet sheet;
    int counter = 1;
    ModbusServer modbusServer = new ModbusServer();
    static Timer myTimer;
    MyTimerTask(Timer timer) {
        myTimer=timer;
        modbusServer.setPort(502);// Note that Standard Port for Modbus TCP communication is 502
        try {
            workbook = new HSSFWorkbook();
            sheet = workbook.createSheet("Sample Sheet");
            HSSFRow row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Date");
            cell = row.createCell(1);
            cell.setCellValue("Value");

            modbusServer.coils[1] = true;
            modbusServer.Listen();
        } catch (IOException e) {

            e.printStackTrace();
        }
    }

    public void run() {
        num++;
        Date date = new Date();
        SimpleDateFormat DateFor = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss.SSS");
        String stringDate = DateFor.format(date);

        System.out.println(" Time = " + stringDate + "    value = " + num);
        modbusServer.holdingRegisters[1] = num;
        modbusServer.holdingRegisters[2] = num;
        modbusServer.holdingRegisters[3] = num;

        HSSFRow row = sheet.createRow(counter);
        Cell cell = row.createCell(0);
        cell.setCellValue(stringDate);
        cell = row.createCell(1);
        cell.setCellValue(num);
        counter++;
        if (counter > 300) {
            try {

                FileOutputStream out = new FileOutputStream(new File("data.xls"));
                workbook.write(out);
                workbook.close();
                System.out.println("kapandı");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            myTimer.cancel();
            return;
        }
    }
}