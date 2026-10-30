import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.opencsv.*;

public class Main {

    private static double[] GenerateRandomNumbers(long seed, int amount) {
        double[] randomList = new double[amount];
        for (int i=0;i<amount;i++) {
            Random generator = new Random(seed);
            randomList[i] = Math.abs((double) (generator.nextLong() % 0.001) * 10000);
            seed--;
        }
        return randomList;

    } //end of method GenerateRandomNumbers()

    private static void TEST_GenerateRandomNumbers() {
        long seedLong = (long)(2.3);
        int amountInt = 10;
        double[] randomListTest = GenerateRandomNumbers(seedLong, amountInt);
        int lengthofrandomListTest = randomListTest.length;
        for (int i=0; i< lengthofrandomListTest; i++) {
            System.out.println("Random number <" + ((i+1)) + "> is : " + randomListTest[i]);
        }

    } //end of method TEST_GenerateRandomNumbers()

    public static void main(String[] args) throws Exception {

        String outputdirectory = "C:\\Temp\\";

        //Prompt user to input values for parameters i,j,k,m
        Scanner scanner1 = new Scanner(System.in);



        System.out.println("==================== DELL INVENTORY REPOSITIONING MODEL ===================");
        System.out.println("=============================================");
        System.out.println("TO EXIT/ABORT DURING DATA ENTRY, HIT Ctrl-D ");
        System.out.println("=============================================");
        System.out.println("The indices i, j, and m are used to represent SITES, and index k is used to represent PARTS.");
        System.out.println("The SITES indexed by i supply to the SITES indexed by j.");
        System.out.println("The SITES indexed by m supply to the SITES indexed by j.");
        System.out.println("The SITES indexed by i and m have no relationships between them.");
        System.out.println("==============================================");
        System.out.println("PLEASE ENTER INTEGER VALUES ONLY.");
        System.out.println("=============================================");
        System.out.println("Enter values for: i<HIT ENTER>, j<HIT ENTER>, m<HIT ENTER>, k<HIT ENTER> : ");
        System.out.println("For example: 5<HIT ENTER>, 5<HIT ENTER>, 3<HIT ENTER>, 1<HIT ENTER>");

        //EXCESS is a random number between 20 and 150
        //NEEDED is a random number between 20 and 150
        //COST_TO_SHIP_FROM_M_TO_J is a random number between 75 and 175
        //COST_TO_SHIP_FROM_I_TO_J is a random number between 50 and 150

        //SITES:
        int para_i=0, para_j=0, para_m=0;

        //PARTS
        int para_k=0;

        try {
            para_i = scanner1.nextInt();
            para_j = scanner1.nextInt();
            para_m = scanner1.nextInt();
            para_k = scanner1.nextInt();

            // Display the values that were entered by the user
            System.out.println("\nYOU ENTERED:\n");
            System.out.println("Variable i (number of i sites) = " + para_i);
            System.out.println("Variable j (number of j sites) = " + para_j);
            System.out.println("Variable m (number of m sites) = " + para_m);
            System.out.println("Variable k (number of parts) = " + para_k);

        }
        catch (InputMismatchException ime) {
            System.out.println("\nEXCEPTION CAUGHT: InputMismatchException. The number you enter must be an integer for each input. Please exit and rerun the application and then reenter the inputs.\n");
            ime.printStackTrace();
        }




        System.out.println("\nNext, please enter two INTEGER numbers: the first number is the number of problem instances that you want to generate and the second number is the instance serial number (ISN) of the first instance.");
        System.out.println("If you need to generate two or more instances, then you would still enter the ISN of the first instance. In this case, the ISNs of all instances after the first instance, will be auto-incremented by 1 after the ISN of the first instance.");
        System.out.println("For example if you enter NUMBER_OF_INSTANCES_TO_BE_GENERATED as 3, and enter INSTANCE_SERIAL_NUMBER_OF_THE_FIRST_INSTANCE as 156, then you will get 3 instances with serial numbers 156, 157, and 158.");
        System.out.println("\nEnter values for: NUMBER_OF_INSTANCES_TO_BE_GENERATED<HIT ENTER>, INSTANCE_SERIAL_NUMBER_OF_THE_FIRST_INSTANCE<HIT ENTER>");

        int numberofinstances=0;
        int ISNofthefirstinstance=0;

        Scanner scanner2 = new Scanner(System.in);

        try {

            numberofinstances = scanner2.nextInt();
            ISNofthefirstinstance = scanner2.nextInt();

            // Display the values that were entered by the user
            System.out.println("\nYOU ENTERED:\n");
            System.out.println("Number of instances to be generated = " + numberofinstances);
            System.out.println("ISN of the first instance = " + ISNofthefirstinstance);
            System.out.println("=========================================\n");
            System.out.println("\n=========================================\n");
            System.out.println("\nThe instances are now being generated...\n");
            System.out.println("\n=========================================\n");
        }
        catch (InputMismatchException ime) {
            System.out.println("\n EXCEPTION CAUGHT: InputMismatchException. The number you enter must be an integer for each input. Please exit and rerun the application and then reenter the inputs.\n");
            ime.printStackTrace();
        }

        //Run TESTER for GenerateRandomNumbers()
        //TEST_GenerateRandomNumbers();




        //Generate instances

        for(int index1 = 0; index1 < numberofinstances; index1++ ) {

            int ISN = ISNofthefirstinstance + index1;
            long seed = (long)ISN;
            Random generator = new Random(seed);

            //Use the string instance number as the <EXCEL Workbook filename>.XSLX for this instance.
            String StringISN = Integer.toString(ISN);


            //Generate worksheets for this instance and write the worksheets to a single workbook for this instance.

            //Generate workbook
            XSSFWorkbook workbook = new XSSFWorkbook();

            //Generate worksheets in this workbook
            XSSFSheet sheet_PARA_Eik = workbook.createSheet("PARA_Eik");
            XSSFSheet sheet_PARA_Njk = workbook.createSheet("PARA_Njk");
            XSSFSheet sheet_PARA_Bmjk = workbook.createSheet("PARA_Bmjk");
            XSSFSheet sheet_PARA_Cijk = workbook.createSheet("PARA_Cijk");

            //Generate rows in each worksheet
            XSSFRow ROW_in_sheet_PARA_Eik, ROW_in_sheet_PARA_Njk, ROW_in_sheet_PARA_Bmjk, ROW_in_sheet_PARA_Cijk;

            int EXCESS = 0;
            int NEEDED = 0;
            int COST_M_TO_J = 0;
            int COST_I_TO_J = 0;

            //Map<Integer, Object[]> DATA_Param_Eik = new TreeMap<Integer, Object[]>();
            Map<Integer, String[]> DATA_Param_Eik = new TreeMap<Integer, String[]>();
            int DATA_Param_Eik_rowid = 0;
            //DATA_Param_Eik.put(DATA_Param_Eik_rowid++, new Object[] {"PART k", "SITE i", "EXCESS UNITS"});
            DATA_Param_Eik.put(DATA_Param_Eik_rowid++, new String[] {"PART k", "SITE i", "EXCESS UNITS"});

            //Map<Integer, Object[]> DATA_Param_Njk = new TreeMap<Integer, Object[]>();
            Map<Integer, String[]> DATA_Param_Njk = new TreeMap<Integer, String[]>();
            int DATA_Param_Njk_rowid = 0;
            //DATA_Param_Njk.put(DATA_Param_Njk_rowid++, new Object[] {"PART k", "SITE j", "NEEDED UNITS"});
            DATA_Param_Njk.put(DATA_Param_Njk_rowid++, new String[] {"PART k", "SITE j", "NEEDED UNITS"});

            if (para_i == para_j) {
                //This means that we can correlate EXCESS and NEEDED

                for(int index2 = 1; index2 <= para_k; index2++) {
                    for (int index3 = 1; index3 <= para_i; index3++) {
                        //Generate worksheet PARA_Eik and PARA_Njk simultaneously

                        //This means that the i-SITES and j-SITES are equal.
                        //Hence, we can enforce mutual exclusivity between NEEDED and EXCESS.

                        int temprandom1 = generator.nextInt(1, 100);
                        if (temprandom1 <= 50) {
                            EXCESS = 0;
                            NEEDED = generator.nextInt(20, 150);
                        } else {
                            EXCESS = generator.nextInt(20, 150);
                            NEEDED = 0;
                        }
                        //DATA_Param_Eik.put(DATA_Param_Eik_rowid++, new Object[] {Integer.toString(index2), Integer.toString(index3), Integer.toString(EXCESS)});
                        //DATA_Param_Njk.put(DATA_Param_Njk_rowid++, new Object[] {Integer.toString(index2), Integer.toString(index3), Integer.toString(NEEDED)});
                        DATA_Param_Eik.put(DATA_Param_Eik_rowid++, new String[] {Integer.toString(index2), Integer.toString(index3), Integer.toString(EXCESS)});
                        DATA_Param_Njk.put(DATA_Param_Njk_rowid++, new String[] {Integer.toString(index2), Integer.toString(index3), Integer.toString(NEEDED)});

                    } //end of FOR loop for PARAM_I and PARAM_J (since they are equal)

                } //end of FOR loop for PARAM_K

            } //end of logic to process correlated EXCESS and NEEDED

            else {
                //This means that we CANNOT correlate EXCESS and NEEDED
                //Set random values for EXCESS and NEEDED independently.

                for(int index4 = 1; index4 <= para_k; index4++) {

                    for (int index5 = 1; index5 <= para_i; index5++) {
                        //Generate worksheet PARA_Eik
                        //First column: Part k
                        //Second column: Site i
                        //Third column: EXCESS UNITS (randomly generated)

                        //This means there are unequal number of i-SITES and j-SITES.
                        //Hence, mutual exclusivity cannot be enforced between EXCESS and NEEDED.
                        //Hence, randomly generate quantities for EXCESS and NEEDED
                        EXCESS = generator.nextInt(20, 150);
                        //DATA_Param_Eik.put(DATA_Param_Eik_rowid++, new Object[] {Integer.toString(index4), Integer.toString(index5), Integer.toString(EXCESS)});
                        DATA_Param_Eik.put(DATA_Param_Eik_rowid++, new String[] {Integer.toString(index4), Integer.toString(index5), Integer.toString(EXCESS)});

                    } //end of FOR loop for PARAM_I

                    for (int index6 = 1; index6 <= para_j; index6++) {
                        //Generate worksheet PARA_Njk
                        //First column: Part k
                        //Second column: Site j
                        //Third column: NEEDED UNITS (randomly generated)
                        NEEDED = generator.nextInt(20, 150);
                        //DATA_Param_Njk.put(DATA_Param_Njk_rowid++, new Object[] {Integer.toString(index4), Integer.toString(index6), Integer.toString(NEEDED)});
                        DATA_Param_Njk.put(DATA_Param_Njk_rowid++, new String[] {Integer.toString(index4), Integer.toString(index6), Integer.toString(NEEDED)});

                    } //end of FOR loop for PARAM_J

                } //end of FOR loop for PARAM_K

            } //end of ELSE ie we cannot correlate EXCESS and NEEDED

            //Save data from DATA_Param_Eik to its own worksheet
            Set<Integer> DATA_Param_Eik_KEYS = DATA_Param_Eik.keySet();
            DATA_Param_Eik_rowid = 0;
            for (Integer key: DATA_Param_Eik_KEYS) {
                ROW_in_sheet_PARA_Eik = sheet_PARA_Eik.createRow(DATA_Param_Eik_rowid++);
                Object[] objectArray = DATA_Param_Eik.get(key);
                int cellid = 0;
                for (Object obj: objectArray) {
                    Cell cell = ROW_in_sheet_PARA_Eik.createCell(cellid++);
                    //The line below works....but it create the green triangle warning in Excel which throws off CPLEX.
                    cell.setCellValue((String)obj);

                    //Hence, I want to try to see if I can save a numeric value in the EXCEL CELL.
                    //The code block below DID NOT WORK.
                    /*
                    CellStyle style = workbook.createCellStyle();
                    DataFormat format = workbook.createDataFormat();
                    String cellformat = "0.00";
                    style.setDataFormat(format.getFormat(cellformat));
                    double tmpvalue = Double.parseDouble((String)obj);
                    cell.setCellValue(tmpvalue);
                    cell.setCellStyle(style);
                    */
                }

            }

            //Save data from DATA_Param_Njk to its own worksheet
            Set<Integer> DATA_Param_Njk_KEYS = DATA_Param_Njk.keySet();
            DATA_Param_Njk_rowid = 0;
            for (Integer key: DATA_Param_Njk_KEYS) {
                ROW_in_sheet_PARA_Njk = sheet_PARA_Njk.createRow(DATA_Param_Njk_rowid++);
                Object[] objectArray = DATA_Param_Njk.get(key);
                int cellid = 0;
                for (Object obj: objectArray) {
                    Cell cell = ROW_in_sheet_PARA_Njk.createCell(cellid++);
                    cell.setCellValue((String)obj);
                }

            }



            //Generate worksheet PARA_Bmjk
            //First column: Part k
            //Second column: Site m
            //Third column: Site j
            //Fourth column: COST m to j (randomly generated)
            //Map<Integer, Object[]> DATA_Param_Bmjk = new TreeMap<Integer, Object[]>();
            Map<Integer, String[]> DATA_Param_Bmjk = new TreeMap<Integer, String[]>();
            int DATA_Param_Bmjk_rowid = 0;
            //DATA_Param_Bmjk.put(DATA_Param_Bmjk_rowid++, new Object[] {"PART k", "SITE m", "SITE j", "COST m to j"});
            DATA_Param_Bmjk.put(DATA_Param_Bmjk_rowid++, new String[] {"PART k", "SITE m", "SITE j", "COST m to j"});

            for (int index7 = 1; index7 <= para_k; index7++) {
                for (int index8 = 1; index8 <= para_m; index8++) {
                    for (int index9 = 1; index9 <= para_j; index9++) {
                        if (index8 == index9) {
                            COST_M_TO_J = 0;
                        }

                        else {
                            COST_M_TO_J = generator.nextInt(75, 175);
                        }

                        //DATA_Param_Bmjk.put(DATA_Param_Bmjk_rowid++, new Object[] {Integer.toString(index7), Integer.toString(index8), Integer.toString(index9), Integer.toString(COST_M_TO_J)});
                        DATA_Param_Bmjk.put(DATA_Param_Bmjk_rowid++, new String[] {Integer.toString(index7), Integer.toString(index8), Integer.toString(index9), Integer.toString(COST_M_TO_J)});


                    }
                }
            }

            //Save data from DATA_Param_Bmjk to its own worksheet
            Set<Integer> DATA_Param_Bmjk_KEYS = DATA_Param_Bmjk.keySet();
            DATA_Param_Bmjk_rowid = 0;
            for (Integer key: DATA_Param_Bmjk_KEYS) {
                ROW_in_sheet_PARA_Bmjk = sheet_PARA_Bmjk.createRow(DATA_Param_Bmjk_rowid++);
                Object[] objectArray = DATA_Param_Bmjk.get(key);
                int cellid = 0;
                for (Object obj: objectArray) {
                    Cell cell = ROW_in_sheet_PARA_Bmjk.createCell(cellid++);
                    cell.setCellValue((String)obj);
                }

            }


            //Generate worksheet PARA_Cijk
            //First column: Part k
            //Second column: Site i
            //Third column: Site j
            //Fourth column: COST i to j (randomly generated)
            //Map<Integer, Object[]> DATA_Param_Cijk = new TreeMap<Integer, Object[]>();
            Map<Integer, String[]> DATA_Param_Cijk = new TreeMap<Integer, String[]>();
            int DATA_Param_Cijk_rowid = 0;
            //DATA_Param_Cijk.put(DATA_Param_Cijk_rowid++, new Object[] {"PART k", "SITE i", "SITE j", "COST i to j"});
            DATA_Param_Cijk.put(DATA_Param_Cijk_rowid++, new String[] {"PART k", "SITE i", "SITE j", "COST i to j"});


            for (int index10 = 1; index10 <= para_k; index10++) {
                for (int index11 = 1; index11 <= para_i; index11++) {
                    for (int index12 = 1; index12 <= para_j; index12++) {
                        if (index11 == index12) {
                            COST_I_TO_J = 0;
                        }

                        else {
                            COST_I_TO_J = generator.nextInt(50, 150);
                        }

                        //DATA_Param_Cijk.put(DATA_Param_Cijk_rowid++, new Object[] {Integer.toString(index10), Integer.toString(index11), Integer.toString(index12), Integer.toString(COST_I_TO_J)});
                        DATA_Param_Cijk.put(DATA_Param_Cijk_rowid++, new String[] {Integer.toString(index10), Integer.toString(index11), Integer.toString(index12), Integer.toString(COST_I_TO_J)});


                    }
                }
            }

            //Save data from DATA_Param_Cijk to its own worksheet
            Set<Integer> DATA_Param_Cijk_KEYS = DATA_Param_Cijk.keySet();
            DATA_Param_Cijk_rowid = 0;
            for (Integer key: DATA_Param_Cijk_KEYS) {
                ROW_in_sheet_PARA_Cijk = sheet_PARA_Cijk.createRow(DATA_Param_Cijk_rowid++);
                Object[] objectArray = DATA_Param_Cijk.get(key);
                int cellid = 0;
                for (Object obj: objectArray) {
                    Cell cell = ROW_in_sheet_PARA_Cijk.createCell(cellid++);
                    cell.setCellValue((String)obj);
                }

            }

            //Write all the four worksheets to a single EXCEL workbook
            String XLSX_filepath = outputdirectory + StringISN + ".xlsx";
            FileOutputStream out = new FileOutputStream(new File(XLSX_filepath));
            workbook.write(out);
            out.close();

            //For each instance, write each worksheet (not workbook) to its own CSV file in its own instance ISN directory
            //For each instance, create a directory C:\Temp\<ISN>
            //Then in each directory, save the CSV files for all the worksheets for this instance.
            //
            //For e.g., If I have 2 instances: 141 and 142 and each instance has 4 files, then I will generate the following structure:
            //C:\Temp\141\141_Param_Eik.csv
            //C:\Temp\141\141_Param_Njk.csv
            //C:\Temp\141\141_Param_Bmjk.csv
            //C:\Temp\141\141_Param_Cijk.csv
            //
            //C:\Temp\142\142_Param_Eik.csv
            //C:\Temp\142\142_Param_Njk.csv
            //C:\Temp\142\142_Param_Bmjk.csv
            //C:\Temp\142\142_Param_Cijk.csv


            String CSV_filepath_for_Param_Eik = outputdirectory + StringISN + "\\" + StringISN + "_Param_Eik" + ".CSV";
            String CSV_filepath_for_Param_Njk = outputdirectory + StringISN + "\\" + StringISN + "_Param_Njk" + ".CSV";
            String CSV_filepath_for_Param_Bmjk = outputdirectory + StringISN + "\\" + StringISN + "_Param_Bmjk" + ".CSV";
            String CSV_filepath_for_Param_Cijk = outputdirectory + StringISN + "\\" + StringISN + "_Param_Cijk" + ".CSV";

            File file_CSV_filepath_for_Param_Eik = new File(CSV_filepath_for_Param_Eik);
            File file_CSV_filepath_for_Param_Njk = new File(CSV_filepath_for_Param_Njk);
            File file_CSV_filepath_for_Param_Bmjk = new File(CSV_filepath_for_Param_Bmjk);
            File file_CSV_filepath_for_Param_Cijk = new File(CSV_filepath_for_Param_Cijk);

            //Create the parent directories if they don't exist already.
            //This is necessary the first time the CSV file is created because the parent directory <ISN> does not exist.
            //For subsequent runs of the same ISN, no new directories are created since they already exist.
            file_CSV_filepath_for_Param_Eik.getParentFile().mkdirs();
            file_CSV_filepath_for_Param_Njk.getParentFile().mkdirs();
            file_CSV_filepath_for_Param_Bmjk.getParentFile().mkdirs();
            file_CSV_filepath_for_Param_Cijk.getParentFile().mkdirs();

            //Write Param_Eik CSV file
            try {
                FileWriter FW_Eik = new FileWriter(file_CSV_filepath_for_Param_Eik);
                CSVWriter CSVWriter_Eik = new CSVWriter(FW_Eik);

                for (Integer key: DATA_Param_Eik_KEYS) {
                    String[] ROW_Eik = DATA_Param_Eik.get(key);
                    CSVWriter_Eik.writeNext(ROW_Eik);
                }
                CSVWriter_Eik.close();
            }
            catch (IOException ioe) {
                System.out.println("EXCEPTION CAUGHT: IOException during the writing of Param_Eik CSV file. Please look for the root cause in the stacktrace below.");
                ioe.printStackTrace();
            }

            //Write Param_Njk CSV file
            try {
                FileWriter FW_Njk = new FileWriter(file_CSV_filepath_for_Param_Njk);
                CSVWriter CSVWriter_Njk = new CSVWriter(FW_Njk);

                for (Integer key: DATA_Param_Njk_KEYS) {
                    String[] ROW_Njk = DATA_Param_Njk.get(key);
                    CSVWriter_Njk.writeNext(ROW_Njk);
                }
                CSVWriter_Njk.close();
            }
            catch (IOException ioe) {
                System.out.println("EXCEPTION CAUGHT: IOException during the writing of Param_Njk CSV file. Please look for the root cause in the stacktrace below.");
                ioe.printStackTrace();
            }

            //Write Param_Bmjk CSV file
            try {
                FileWriter FW_Bmjk = new FileWriter(file_CSV_filepath_for_Param_Bmjk);
                CSVWriter CSVWriter_Bmjk = new CSVWriter(FW_Bmjk);

                for (Integer key: DATA_Param_Bmjk_KEYS) {
                    String[] ROW_Bmjk = DATA_Param_Bmjk.get(key);
                    CSVWriter_Bmjk.writeNext(ROW_Bmjk);
                }
                CSVWriter_Bmjk.close();
            }
            catch (IOException ioe) {
                System.out.println("EXCEPTION CAUGHT: IOException during the writing of Param_Bmjk CSV file. Please look for the root cause in the stacktrace below.");
                ioe.printStackTrace();
            }

            //Write Param_Cijk CSV file
            try {
                FileWriter FW_Cijk = new FileWriter(file_CSV_filepath_for_Param_Cijk);
                CSVWriter CSVWriter_Cijk = new CSVWriter(FW_Cijk);

                for (Integer key: DATA_Param_Cijk_KEYS) {
                    String[] ROW_Cijk = DATA_Param_Cijk.get(key);
                    CSVWriter_Cijk.writeNext(ROW_Cijk);
                }
                CSVWriter_Cijk.close();
            }
            catch (IOException ioe) {
                System.out.println("EXCEPTION CAUGHT: IOException during the writing of Param_Cijk CSV file. Please look for the root cause in the stacktrace below.");
                ioe.printStackTrace();
            }


        } //end of FOR loop for each instance


        scanner1.close();
        scanner2.close();




    } //end of main method main()

} //end of class Main



//Notes:
//You can extract different data types from the user input:
//To read a boolean value from the user, use nextBoolean()
//To read a byte value from the user, use nextByte()
//To read a double value from the user, use nextDouble()
//To read a float value from the user, use nextFloat()
//To read an int value from the user, use nextInt()
//To read a long value from the user, use nextLong()
//To read a short value from the user, use nextShort()
//To read a string value from the user, use nextLine()

//String conversions:
//Note: new Integer(String) to convert String to Integer is DEPRECATED.
//Use either of these two methods instead:
//Integer.parseInt(String) converts String to int
//Integer.valueOf(String) converts String to Integer