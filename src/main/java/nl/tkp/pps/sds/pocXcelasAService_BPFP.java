/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package nl.tkp.pps.sds;

import java.io.FileInputStream;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

/**
 * Attempts to re-evaluate all the formulas in the workbook, and
 *  reports what (if any) formula functions used are not (currently)
 *  supported by Apache POI.
 *
 * <p>This provides examples of how to evaluate formulas in excel
 *  files using Apache POI, along with how to handle errors whilst
 *  doing so.
 */


// 20170615_PPS-WOD_KPN.xlsm

public class pocXcelasAService_BPFP {

    static FormulaEvaluator evaluator = null;

    public static void main(String[] args) throws Exception {
        
        String fname = new String("./files/PPS-WON_BPF_Particuliere_Beveiliging_v2.0.xlsx");
        String mainSheetNaam = new String("BPF_Particuliere Beveiliging");
        String invoerSheetNaam = new String("Invoer_BPFP Beveiliging");

        Map inv = null;

        Invoer_BPFP i_1 = new Invoer_BPFP.InvoerBuilder(1234,"M", new DateTime("1959-12-13"))
                .rekenDatum( new DateTime("2017-01-02")).betaalDatum(new DateTime("2017-02-01"))
                .uurLoon(38.50)
                .gewerkteUrenEersteVolledigePeriode(140)
                .normUrenPerWeek(40)
                .verloningsPeriode("Maand")
                .maandEersteLoongegevens("Januari")
                .overdrachtsWaardeOP(84000)
                .overdrachtsWaardeBPP(0)
                .overdrachtsWaardeVerevendOP(15000)
                .werkgeversNummer(1234)
                .circuit("SDS")
                .pensioenLeeftijd(67)
                .build();

        Invoer_BPFP i_2 = new Invoer_BPFP.InvoerBuilder(6789,"V", new DateTime("1969-12-13"))
                .rekenDatum( new DateTime("2018-02-01")).betaalDatum(new DateTime("2018-04-01"))
                .uurLoon(45.65)
                .gewerkteUrenEersteVolledigePeriode(131)
                .normUrenPerWeek(39)
                .verloningsPeriode("Maand")
                .maandEersteLoongegevens("Maart")
                .overdrachtsWaardeOP(32768)
                .overdrachtsWaardeBPP(456)
                .overdrachtsWaardeVerevendOP(13456)
                .werkgeversNummer(4056000)
                .circuit("SDS")
                .pensioenLeeftijd(67)
                .build();



        FileInputStream fis = new FileInputStream(fname);
        Workbook wb = new XSSFWorkbook(fis); 
        Sheet mainSheet = wb.getSheet(mainSheetNaam);
        Sheet invoerSheet = wb.getSheet(invoerSheetNaam);
        evaluator = wb.getCreationHelper().createFormulaEvaluator();

        Map<String,String> uitv = getUitvoer(mainSheet,new pocRange2D("I2:J16") );

        // Dit is de setter om invoer door te laten rekenen
        setSingleCellValue(invoerSheet, new pocRange1D("D2"),99);

        // Test 1
        insertInvoer(invoerSheet, "22", i_1);
        inv = getInvoer( invoerSheet, 99);
        uitv = getUitvoer(mainSheet,new pocRange2D("I2:J16") );
        printMap("Invoer_BPFP-99-i_2"," ",inv);
        printMap("Uitvoer-99-i_2"," ",uitv);


        // Test 2
        insertInvoer(invoerSheet, "22", i_2);
        inv = getInvoer( invoerSheet, 99);
        uitv = getUitvoer(mainSheet,new pocRange2D("I2:J16") );
        printMap("Invoer_BPFP-99-i_2"," ",inv);
        printMap("Uitvoer-99-i_2"," ",uitv);


        //FileOutputStream fileOut = new FileOutputStream("saved.xlsx");
        //wb.write(fileOut);
        //fileOut.close();



    }

    private static void insertInvoer(Sheet sht, String startRow, Invoer_BPFP inv){

        setSingleCellValue(sht, new pocRange1D("B"+startRow), inv.pensioenNummer);
        setSingleCellValue(sht, new pocRange1D("C"+startRow), inv.geslacht);
        setSingleCellValue(sht, new pocRange1D("D"+startRow), inv.geboorteDatum);
        setSingleCellValue(sht, new pocRange1D("E"+startRow), inv.rekenDatum);
        setSingleCellValue(sht, new pocRange1D("F"+startRow), inv.betaalDatum);
        setSingleCellValue(sht, new pocRange1D("G"+startRow), inv.uurLoon);
        setSingleCellValue(sht, new pocRange1D("H"+startRow), inv.gewerkteUrenEersteVolledigePeriode);
        setSingleCellValue(sht, new pocRange1D("I"+startRow), inv.normUrenPerWeek);
        setSingleCellValue(sht, new pocRange1D("J"+startRow), inv.verloningsPeriode);
        setSingleCellValue(sht, new pocRange1D("K"+startRow), inv.maandEersteLoongegevens);
        setSingleCellValue(sht, new pocRange1D("L"+startRow), inv.overdrachtsWaardeOP);
        setSingleCellValue(sht, new pocRange1D("M"+startRow), inv.overdrachtsWaardeBPP);
        setSingleCellValue(sht, new pocRange1D("N"+startRow), inv.overdrachtsWaardeVerevendOP);
        setSingleCellValue(sht, new pocRange1D("O"+startRow), inv.werkgeversNummer);
        setSingleCellValue(sht, new pocRange1D("P"+startRow), inv.circuit);
        setSingleCellValue(sht, new pocRange1D("Q"+startRow), inv.pensioenLeeftijd);


    }

    private static void printMap(String header, String prefix, Map<String,String> hm){
        System.out.println(header);

        Iterator<Map.Entry<String, String>> entries = hm.entrySet().iterator();
        while (entries.hasNext()) {
            Map.Entry<String,String> entry = entries.next();
            System.out.println(prefix+" => tag=" + entry.getKey() + " value = " + entry.getValue());
        }


    }

    private static String CellValueAsString( Cell c){
        //System.out.println("T="+c.getCellTypeEnum());
        DataFormatter df = new DataFormatter();
        if(c == null)return null;

        CellType cellType = c.getCellTypeEnum();
        switch(cellType) {
            case NUMERIC:
                if( DateUtil.isCellDateFormatted(c)){
                    org.joda.time.DateTime jodaDT = new org.joda.time.DateTime( c.getDateCellValue() );
                    return jodaDT.getYear()+"-"+jodaDT.getMonthOfYear()+"-"+jodaDT.getDayOfMonth();
                }
                else {
                    //return df.formatCellValue(c);
                    return Double.toString(c.getNumericCellValue());
                }
            case ERROR:
                return "";

            case STRING:
                return c.getStringCellValue();
            case FORMULA:
                CellValue cv = evaluator.evaluate(c);
                switch(cv.getCellTypeEnum()){
                    case NUMERIC:
                        return Double.toString( cv.getNumberValue());
                    case ERROR:
                        return "";
                    case STRING:
                        return cv.getStringValue();
                    default:
                        return "error formula";
                }
            default:
                    return null;
        }
    }

    public static Map<String, String> getUitvoer(Sheet sht, pocRange2D range ){

        HashMap hm = new HashMap<String,String>();

        for(int row=range.startRow; row<range.endRow; row++){
            Row r = sht.getRow(row);
            String tag = CellValueAsString(r.getCell(range.startColumn) );
            String val = CellValueAsString( r.getCell(range.endColumn) );

            if(tag != null && tag.length()!=0){
                hm.put(tag,val);
            }
        }

        return hm;
    }

    public static String getSingleCellValue(Sheet sht, pocRange1D range){

        Row r = sht.getRow( range.row );
        String val = CellValueAsString( r.getCell( range.column));

        return val;
    }

    private static Cell getSingleCell(Sheet sht, pocRange1D range){
        Row r = sht.getRow( range.row);
        Cell c = r.getCell( range.column);
        return c;
    }

    public static void setSingleCellValue(Sheet sht, pocRange1D range, int val){
        Cell c = getSingleCell(sht,range);
        c.setCellValue( val );
        evaluator.notifyUpdateCell(c);
    }

    public static void setSingleCellValue(Sheet sht, pocRange1D range, DateTime val){
        Cell c = getSingleCell(sht,range);
        String s = c.getCellStyle().getDataFormatString();

        boolean b = DateUtil.isCellDateFormatted(c);

        Date d = val.toDate();
        double excel_d = DateUtil.getExcelDate(d);

        c.setCellValue(d);
        evaluator.notifyUpdateCell(c);

    }

    public static void setSingleCellValue(Sheet sht, pocRange1D range, double val){
        Cell c = getSingleCell(sht,range);
        c.setCellValue(val);
        evaluator.notifyUpdateCell(c);

    }

    public static void setSingleCellValue(Sheet sht, pocRange1D range, String val){
        Cell c = getSingleCell(sht,range);
        c.setCellValue(val);
        evaluator.notifyUpdateCell(c);
    }


    public static Map<String, String> getInvoer(Sheet invoer_sht, int invoerNr ){

        HashMap hm = new HashMap<String,String>();

        Row invoerRow=null;


        // Zoek juiste row bij invoerNr. Invoer_BPFP nr zit in col A (=0)

        for(int i=2;i<30;i++){
            Row r = invoer_sht.getRow(i);
            //System.out.println("cv="+i+" "+CellValueAsString( r.getCell(0)));
            Cell t = r.getCell(0);
            String cellValue = CellValueAsString(r.getCell(0));

            if(  cellValue != null && cellValue.length()>0 && cellValue.compareTo(Integer.toString(invoerNr)+".0" ) == 0){
                invoerRow = r;
                break;
            }

        }

        if(invoerRow==null){
            System.out.println("Geen invoer gevonden voor "+invoerNr);
            System.exit(1);
        }

        Row tagrow = invoer_sht.getRow(1);
        Row valrow = invoerRow;

        for( int col=0;col<25;col++){
            String valString = CellValueAsString( valrow.getCell(col));
            //System.out.println("Fv "+col+"=>"+valString + " "+ valrow.getCell(col).getAddress().formatAsString() );
            if( valString != null && valString.length() > 0){
                String tagString = CellValueAsString( tagrow.getCell(col));
                //System.out.println("Tag ="+tagString);
                hm.put(tagString,valString);
            }
        }

        return hm;
    }

}