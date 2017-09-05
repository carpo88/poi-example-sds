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

import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.udf.AggregatingUDFFinder;
import org.apache.poi.ss.formula.udf.DefaultUDFFinder;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import java.io.FileInputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Attempts to re-evaluate all the formulas in the workbook, and
 *  reports what (if any) formula functions used are not (currently)
 *  supported by Apache POI.
 *
 * <p>This provides examples of how to evaluate formulas in excel
 *  files using Apache POI, along with how to handle errors whilst
 *  doing so.
 */



public class pocXcelasAService_PFKPN55MIN {

    static FormulaEvaluator evaluator = null;

    public static void main(String[] args) throws Exception {

        String fname = new String("./files/20170615_PPS-WOD_KPN.xlsm");
        String mainSheetNaam = new String("KPN55min");
        String invoerSheetNaam = new String("Invoer");
        Map inv = null;
        Map<String,String> uitv = null;

        //
        String[] functionNames = { "tarief" } ;
        FreeRefFunction[] functionImpls = { new udf_PFKPN55MIN_tarief() } ;

        UDFFinder udfs = new DefaultUDFFinder( functionNames, functionImpls ) ;
        UDFFinder udfToolpack = new AggregatingUDFFinder( udfs ) ;


        //

        Invoer_PFKPN55MIN i_1 = new Invoer_PFKPN55MIN.InvoerBuilder()
                .invoernr(99)
                .geslacht("V")
                .burgerlijkeStaat("gehuwd")
                .geboorteDatum( new DateTime("1980-6-16"))
                .pensioenDatum( new DateTime("2047-06-01"))
                .uitdienstDatum( new DateTime("2012-11-15"))
                .rekenDatum( new DateTime("2014-06-16"))
                .betaalDatum( new DateTime("2014-09-01"))
                .regeling("STPS")
                .depot("C")
                .LOP67_excl_verevening(1000)
                .verevend_LOP67(1000)
                .OPO(100)
                .PP_middelloon_excl_BPP(200)
                .PP_bijzonder_LOP(100)
                .PP_eindloon_oud_excl_BPP(300)
                .PP_bijzonder_oud(150)
                .BPR(10000)
                .IPS(2000)
                .PPS(3000)
                .build();



        FileInputStream fis = new FileInputStream(fname);
        Workbook wb = new XSSFWorkbook(fis);
        Sheet mainSheet = wb.getSheet(mainSheetNaam);
        Sheet invoerSheet = wb.getSheet(invoerSheetNaam);

        wb.addToolPack(udfToolpack);


        evaluator = wb.getCreationHelper().createFormulaEvaluator();

        setSingleCellValue(mainSheet, new pocRange1D("C1"),1);
        inv = getInvoer( invoerSheet, 1);
        uitv = getUitvoer(mainSheet,new pocRange2D("A34:C42") );
        printMap("Invoer_BPFP-1"," ",inv);
        printMap("Uitvoer-1"," ",uitv);


        // Dit is de setter om invoer door te laten rekenen
        setSingleCellValue(mainSheet, new pocRange1D("C1"),99);

        // Test 1
        insertInvoer(invoerSheet, "22", i_1);
        inv = getInvoer( invoerSheet, 99);
        uitv = getUitvoer(mainSheet,new pocRange2D("A34:C42") );
        printMap("Invoer_BPFP-99-i_2"," ",inv);
        printMap("Uitvoer-99-i_2"," ",uitv);



        //FileOutputStream fileOut = new FileOutputStream("saved.xlsx");
        //wb.write(fileOut);
        //fileOut.close();



    }

    private static void insertInvoer(Sheet sht, String startRow, Invoer_PFKPN55MIN inv){

        setSingleCellValue(sht, new pocRange1D("A"+startRow), inv.invoernr);
        setSingleCellValue(sht, new pocRange1D("B"+startRow), inv.geslacht);
        setSingleCellValue(sht, new pocRange1D("C"+startRow), inv.burgerlijkeStaat);
        setSingleCellValue(sht, new pocRange1D("D"+startRow), inv.geboorteDatum);
        setSingleCellValue(sht, new pocRange1D("E"+startRow), inv.pensioenDatum);
        setSingleCellValue(sht, new pocRange1D("F"+startRow), inv.uitdienstDatum);
        setSingleCellValue(sht, new pocRange1D("G"+startRow), inv.rekenDatum);
        setSingleCellValue(sht, new pocRange1D("H"+startRow), inv.betaalDatum);
        setSingleCellValue(sht, new pocRange1D("I"+startRow), inv.regeling);
        setSingleCellValue(sht, new pocRange1D("J"+startRow), inv.depot);
        setSingleCellValue(sht, new pocRange1D("K"+startRow), inv.LOP67_excl_verevening);
        setSingleCellValue(sht, new pocRange1D("L"+startRow), inv.verevend_LOP67);
        setSingleCellValue(sht, new pocRange1D("M"+startRow), inv.OPO);
        setSingleCellValue(sht, new pocRange1D("N"+startRow), inv.PP_middelloon_excl_BPP);
        setSingleCellValue(sht, new pocRange1D("O"+startRow), inv.PP_bijzonder_LOP);
        setSingleCellValue(sht, new pocRange1D("P"+startRow), inv.PP_eindloon_oud_excl_BPP);
        setSingleCellValue(sht, new pocRange1D("Q"+startRow), inv.PP_bijzonder_oud);
        setSingleCellValue(sht, new pocRange1D("R"+startRow), inv.BPR);
        setSingleCellValue(sht, new pocRange1D("S"+startRow), inv.IPS);
        setSingleCellValue(sht, new pocRange1D("T"+startRow), inv.PPS);


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
                    DateTime jodaDT = new DateTime( c.getDateCellValue() );
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

    private static Cell createSingleCell(Sheet sht, pocRange1D range, String cellStringType){
        Row r = sht.getRow( range.row);
        CellType newCellType=null;
        Cell cell=null;

         if (cellStringType.compareTo("date")==0){

            CreationHelper createHelper = sht.getWorkbook().getCreationHelper();
            Row row = sht.getRow(range.row);


            CellStyle cellStyle = sht.getWorkbook().createCellStyle();
            cellStyle.setDataFormat(
                    createHelper.createDataFormat().getFormat("yyyy-mm-dd"));
            cell = row.createCell(range.column);
            cell.setCellValue(new Date());
            cell.setCellStyle(cellStyle);

        }
        else {
             Row row = sht.getRow(range.row);
             cell = row.createCell(range.column);
         }

        Cell c = r.getCell( range.column);
        return c;
    }

    public static void setSingleCellValue(Sheet sht, pocRange1D range, int val){
        Cell c = getSingleCell(sht,range);
        if(c==null)c=createSingleCell(sht,range,"int");


        c.setCellValue( val );
        evaluator.notifyUpdateCell(c);
    }

    public static void setSingleCellValue(Sheet sht, pocRange1D range, DateTime val){
        Cell c = getSingleCell(sht,range);
        if(c==null)c=createSingleCell(sht,range, "date");

        String s = c.getCellStyle().getDataFormatString();

        boolean b = DateUtil.isCellDateFormatted(c);

        Date d = val.toDate();
        double excel_d = DateUtil.getExcelDate(d);

        c.setCellValue(d);
        evaluator.notifyUpdateCell(c);

    }

    public static void setSingleCellValue(Sheet sht, pocRange1D range, double val){
        Cell c = getSingleCell(sht,range);
        if(c==null)c=createSingleCell(sht,range,"number");

        c.setCellValue(val);
        evaluator.notifyUpdateCell(c);

    }

    public static void setSingleCellValue(Sheet sht, pocRange1D range, String val){
        Cell c = getSingleCell(sht,range);
        if(c==null)c=createSingleCell(sht,range,"string");

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

        Row tagrow = invoer_sht.getRow(2);
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


/*
Function tarief(leeftijd, tabel, kolomnr) As Double
'
'Deze functie kan interpoleren in een (willekeurige)
'tabel, dus bv: 1,1
'               2,1
'etc

Dim x1 As Double
Dim x2 As Double
Dim n As Double

Dim r As Double
Dim tar1 As Double
Dim tar2 As Double

n = 1 - VindRest(tabel, leeftijd)
x1 = Application.RoundDown(leeftijd + n, 0) - n
x2 = Application.RoundUp(leeftijd + n, 0) - n
r = leeftijd - x1
tar1 = Application.VLookup(x1, tabel, kolomnr, False)
tar2 = Application.VLookup(x2, tabel, kolomnr, False)
tarief = (1 - r) * tar1 + r * tar2
End Function

Function VindRest(tabel, leeftijd) As Double
Dim hulp As Double

hulp = Application.VLookup(leeftijd, tabel, 1)

hulp = hulp - Application.RoundDown(hulp, 0)
VindRest = hulp
End Function

Function tarief2(leeftijd, tabel, kolomnr) As Double
'
'Deze functie interpoleert alleen op halve leeftijden
'
Dim x1 As Double
Dim x2 As Double
Dim r As Double
Dim tar1 As Double
Dim tar2 As Double

x1 = Application.RoundDown(leeftijd + 0.5, 0) - 0.5
x2 = Application.RoundUp(leeftijd + 0.5, 0) - 0.5
r = leeftijd - x1
tar1 = Application.VLookup(x1, tabel, kolomnr, False)
tar2 = Application.VLookup(x2, tabel, kolomnr, False)
tarief2 = (1 - r) * tar1 + r * tar2
End Function

Function tariefSamos(leeftijd, plftd, tabel, kolomnr) As Double
'
'Samos interpolatie
'
Dim x1 As Double
Dim x2 As Double
Dim r As Double
Dim tar1 As Double
Dim tar2 As Double

x1 = Application.RoundDown(leeftijd, 0)
x2 = Application.RoundUp(leeftijd, 0)
r = leeftijd - 1
tar1 = Application.VLookup(x1, tabel, kolomnr, False)
tar2 = Application.VLookup(x2, tabel, kolomnr, False)
tariefSamos = ((1 - (leeftijd - x1)) * (plftd - x1) * tar1 _
+ ((leeftijd - x1) * (plftd - 1 - x1) * tar2)) / (plftd - leeftijd)
End Function

Function Rond(a, b) As Double
Dim rest As Double
Dim geheel As Integer
Dim fractie As Double
'Rest = 0, is naar beneden afronden

rest = ((10000 * a) Mod (10000 * b)) / 10000
geheel = Int(a / b)
fractie = Int(2 * rest / b)

Rond = (geheel + fractie) * b

End Function


 */