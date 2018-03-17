package sannaholistichouse.horaastrology;

import android.content.res.AssetManager;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.io.PushbackInputStream;
import java.security.GeneralSecurityException;
import java.util.Iterator;


public class MainActivity extends AppCompatActivity {



    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);



        final Button button = findViewById(R.id.button2);

        button.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                // Code here executes on main thread after user presses button
                checkFields();
                generateChart();
            }
        });
    }

    public void checkFields(){
        //check if input is correct, catch any errors
    }

    public String[][] retrieveFields(){
        String[][] contents = null;
        Password pass = new Password();
        try {
            AssetManager am=getAssets();
            String filename = "1942-43.xlsx";
            InputStream is=am.open("data/ephemeris1942-2040/"+filename);

            Workbook wb = null;

            // Checking of .xlsx file with password protected.
            String isWorkbookLock = "";
            if (!is.markSupported()) {
                is = new PushbackInputStream(is, 8);
            }

            if (POIFSFileSystem.hasPOIFSHeader(is)) {
                POIFSFileSystem fs = new POIFSFileSystem(is);
                EncryptionInfo info = new EncryptionInfo(fs);
                Decryptor d = Decryptor.getInstance(info);
                try {
                    d.verifyPassword(pass.getPassword());
                    is = d.getDataStream(fs);
                    wb = new XSSFWorkbook(OPCPackage.open(is));
                    isWorkbookLock = "true";
                } catch (GeneralSecurityException e) {
                    e.printStackTrace();
                }
            }
            if (isWorkbookLock != "true") {
                wb = new XSSFWorkbook(is);
            }

            // Get contents from the first sheet only
            Sheet s=wb.getSheetAt(0);
            int totalRows=s.getLastRowNum();
//            fieldList.add("rows:" + totalRows);
//            fieldList.add("cols:??");

            contents = new String[totalRows][42];
            Row row = null;

            Iterator rows = s.rowIterator();
            int rowCount = 0;
            while (rows.hasNext()){
                row=(XSSFRow) rows.next();

                int colCount = 0;
                while (colCount < 42){
                    Cell cell = row.getCell(colCount);
                    String content = cell == null ? "" : cell.toString().trim();
                    contents[rowCount][colCount]=content;
                    colCount++;
                }
                rowCount++;
            }

        }catch (Exception e){
            e.printStackTrace();
        }
        return contents;
    }

    public void generateChart(){
        //retrieves from excel, and display it
        String[][] contents = retrieveFields();

//        String msg = "test - ";
//        printCells(contents);

        String[] houseDisplay = new String[12];
        houseDisplay = new String[]{"6","4","2","1","4","6","3","5","7/8","","5","3"};
        display(houseDisplay);
    }

    public void display(String[] houseDisplay){
        final TextView textView1 = (TextView) findViewById(R.id.house1);
        final TextView textView2 = (TextView) findViewById(R.id.house2);
        final TextView textView3 = (TextView) findViewById(R.id.house3);
        final TextView textView4 = (TextView) findViewById(R.id.house4);
        final TextView textView5 = (TextView) findViewById(R.id.house5);
        final TextView textView6 = (TextView) findViewById(R.id.house6);
        final TextView textView7 = (TextView) findViewById(R.id.house7);
        final TextView textView8 = (TextView) findViewById(R.id.house8);
        final TextView textView9 = (TextView) findViewById(R.id.house9);
        final TextView textView10 = (TextView) findViewById(R.id.house10);
        final TextView textView11 = (TextView) findViewById(R.id.house11);
        final TextView textView12 = (TextView) findViewById(R.id.house12);

        textView1.setText(houseDisplay[0]);
        textView2.setText(houseDisplay[1]);
        textView3.setText(houseDisplay[2]);
        textView4.setText(houseDisplay[3]);
        textView5.setText(houseDisplay[4]);
        textView6.setText(houseDisplay[5]);
        textView7.setText(houseDisplay[6]);
        textView8.setText(houseDisplay[7]);
        textView9.setText(houseDisplay[8]);
        textView10.setText(houseDisplay[9]);
        textView11.setText(houseDisplay[10]);
        textView12.setText(houseDisplay[11]);
    }

    public void printCells(String[][] contents){
        System.out.println("START");
        for(String row[] : contents){
            for(String cell : row){
                System.out.print(cell+",");
            }
            System.out.println();
        }
    }
}
