import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Random;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author chimpo
 */
public class GenerateData {

    int howMany;
    String fileName;
    XSSFWorkbook newWorkbook;
    XSSFSheet newSheet;

    Map<String, Object[]> inputData;
    Set < String > keyid;

    String EmpName[] ={
    "Sophia","Emma","Olivia","Isabella","Ava","Lily","Zoe","Chloe","Mia","Madison","Emily","Ella","Madelyn",
    "Abigail","Aubrey","Addison","Avery","Lay", "Hailey","Amelia","Hannah","Charlotte","Kaitlyn","Harper","Kaylee",
    "Sophie","Mackenzie","Peyton","Riley","Grace","Brooklyn","Sarah","Aaliyah","Anna","Arianna","Ellie","Natalie",
    "Isabell","Lillian","Evelyn","Elizabeth","Lyla","Lucy","Claire","Makayla","Kylie","Audrey","Maya","Leah","Gabriella",
    "Annabelle","Savannah","Nora","Reagan","Scarlett","Samantha","Alyssa","Allison","Elena","Stella","Alexis","Victoria",
    "Aria","Molly","Maria","Bailey","Sydney","Bella","Mila","Taylor","Kayla","Eva","Jasmine","Gianna","Alexandra","Julia",
    "Eliana","Kennedy","Brianna","Ruby","Lauren","Alice","Violet","Kendall","Morgan","Caroline","Piper","Brooke","Elise","Alexa",
    "Sienna","Reese","Clara","Paige","Kate","Nevaeh","Sadie","Quinn","David","Anthony","Christian","Colton","Thomas","Dominic",
    "Austin","John","Sebastian","Cooper","Levi","Parker","Isaiah","Chase","Blake","Aaron","Alex","Adam","Tristan",
    "Julian","Jonathan","Christopher","Jace","Nolan","Miles","Jordan","Carson","Colin","Ian","Riley","Xavier","Hudson",
    "Adrian","Cole","Brody","Leo","Jake","Bentley","Sean","Jeremiah","Asher","Nathaniel","Micah","Jason","Ryder","Declan",
    "Hayden","Brandon","Easton","Lincoln","Harrison"
    };
    
    String EmpSurname[]={
    "SMITH","BROWN","JOHNSON","JONES","WILLIAMS","DAVIS","MILLER","WILSON","TAYLOR","CLARK","WHITE","MOORE","THOMPSON","ALLEN",
    "MARTIN","HALL","ADAMS","THOMAS","WRIGHT","BAKER","WALKER","ANDERSON","LEWIS","HARRIS","HILL","KING","JACKSON","LEE","GREEN","WOOD",
    "PARKER","CAMPBELL","YOUNG","ROBINSON","STEWART","SCOTT","ROGERS","ROBERTS","COOK","PHILLIPS","TURNER","CARTER","WARD","FOSTER","MORGAN",
    "HOWARD","COX","JR","BAILEY","RICHARDSON","REED","RUSSELL","EDWARDS","MORRIS","WELLS","PALMER","ANN","MITCHELL","EVANS","GRAY","WHEELER",
    "WARREN","COOPER","BELL","COLLINS","CARPENTER","STONE","COLE","ELLIS","BENNETT","HARRISON","FISHER","HENRY","SPENCER","WATSON","PORTER",
    "NELSON","JAMES","MARSHALL","BUTLER","HAMILTON","TUCKER","STEVENS","WEBB","MAY","WEST","REYNOLDS","HUNT","BARNES","PERKINS","BROOKS",
    "LONG","PRICE","FULLER","POWELL","PERRY","ALEXANDER","RICE","HART","ROSS","ARNOLD","SHAW","FORD","PIERCE","LAWRENCE","HENDERSON","FREEMAN",
    "MASON","ANDREWS","GRAHAM","CHAPMAN","HUGHES","MILLS","GARDNER","JORDAN","BALL","NICHOLS","GIBSON","GREENE","WALLACE","BALDWIN","DAY",
    "DEAVER","SHERMAN","MURPHY","LANE","KNIGHT","HOLMES","BISHOP","KELLY","FRENCH","MYERS","MENTIONED","PATTERSON","ELIZABETH","CASE","CURTIS",
    "SIMMONS","JENKINS","BERRY","HOPKINS","CLARKE","FOX","GORDON","HUNTER","ROBERTSON","PAYNE","JOHNSTON","BARKER","GRANT"
    };
            
    String EmpPet[]= { "Cat" , "Dog" , "Parrot" , "Gozilla", "Monkey", "Ant" , "None"};
    
    public GenerateData(int howMany, String fileName) throws Exception {
        this.howMany = howMany;
        this.fileName = fileName;

        createTabel();
    }

    private GenerateData(){}

    public void openExcel()
    {
        XSSFWorkbook newWorkbook = new XSSFWorkbook();
        XSSFSheet newSheet = newWorkbook.createSheet("EmpData");
    }

    public void genRandomData()
    {
        inputData = new TreeMap<>();

        String EmpId;
        String EmpSalary ;
        String EmpAge;

        for(int row=0; row<howMany; row++)
        {
            String stringRow= Integer.toString(row);

            if(row == 0)
            {
                inputData.put(stringRow,new Object[]{"EmpID","EmpName","EmpSurname","EmpAge","EmpSalary","EmpPet"});
            }
            else
            {
                EmpId= Integer.toString(new Random().nextInt(howMany*2));
                EmpSalary = Integer.toString(new Random().nextInt(10000)+1600);
                EmpAge = Integer.toString(new Random().nextInt(60) + 20);

                inputData.put(stringRow,new Object[]{ EmpId, EmpName[new Random().nextInt(EmpName.length-1)] ,
                        EmpSurname[new Random().nextInt(EmpSurname.length -1)],
                        EmpAge,EmpSalary,EmpPet[new Random().nextInt(EmpPet.length -1)] });

            }

        }
    }

    public void exportDataToExcel() throws IOException
    {
        XSSFRow row;

        Set < String > keyid = inputData.keySet();
        int rowid = 0;
        for (String key : keyid)
        {
            row = newSheet.createRow(rowid++);
            Object [] objectArr = inputData.get(key);
            int cellid = 0;
            for (Object obj : objectArr)
            {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue(obj.toString());
            }
        }

        FileOutputStream fos = new FileOutputStream(fileName+".xlsx");
        newWorkbook.write(fos);
        fos.close();
        newWorkbook.close();
    }

    public void createTabel() throws IOException
    {
        openExcel();
        genRandomData();
        exportDataToExcel();
    }

    public static void main(String[] args) {

        String fileName= JOptionPane.showInputDialog(null,"Add file name:");
        int howMany= Integer.parseInt(JOptionPane.showInputDialog(null,"Add number of rows:"));
       
        try {
            new GenerateData(howMany, fileName);
        } catch (Exception ex) {
            Logger.getLogger(GenerateData.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}
