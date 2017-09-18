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

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.eval.*;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.functions.Function;
import org.apache.poi.ss.formula.udf.AggregatingUDFFinder;
import org.apache.poi.ss.formula.udf.DefaultUDFFinder;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.POILogger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.joda.time.Months;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Attempts to re-evaluate all the formulas in the workbook, and
 *  reports what (if any) formula functions used are not (currently)
 *  supported by Apache POI.
 *
 * <p>This provides examples of how to evaluate formulas in excel
 *  files using Apache POI, along with how to handle errors whilst
 *  doing so.
 */



public class pocPOIKrokus {

    static FormulaEvaluator evaluator = null;


    static HashMap<String, HashMap<Integer,String>> codeTabellen = new HashMap<String, HashMap<Integer, String>>();



    public static void main(String[] args) throws Exception {

        String fname = new String("./files/rekensheet-ingevuld.xlsm");

        String[] functionNames = { "HaalWaardeCodeTabelAlsInt","HaalWaardeCodeTabelAlsFloat" } ;
        FreeRefFunction[] functionImpls = { new udf_HWC(), new udf_HWC() } ;

        UDFFinder udfs = new DefaultUDFFinder( functionNames, functionImpls ) ;
        UDFFinder udfToolpack = new AggregatingUDFFinder( udfs ) ;

        FileInputStream fis = new FileInputStream(fname);

        System.setProperty("POI.FormulaEval", "org.apache.poi.util.SystemOutLogger");
        System.setProperty("org.apache.poi.util.POILogger", "org.apache.poi.util.SystemOutLogger");
        //System.setProperty("poi.log.level", POILogger.DEBUG + "");



        Workbook wb = new XSSFWorkbook(fis);
        evaluator = wb.getCreationHelper().createFormulaEvaluator();


        wb.addToolPack(udfToolpack);


        FunctionEval.registerFunction("DATEDIF", new Function() {
            @Override
            public ValueEval evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
//                System.out.println("DateDiff Call --> "+args+ " "+ srcRowIndex+ "  "+srcColumnIndex);


                ValueEval v_a=null;
                ValueEval v_b=null;
                try{

                    v_a = OperandResolver.getSingleValue(args[0], srcRowIndex,srcColumnIndex);
                    Date d1 = DateUtil.getJavaDate(OperandResolver.coerceValueToDouble(v_a));

                    v_b = OperandResolver.getSingleValue(args[1], srcRowIndex,srcColumnIndex);
                    Date d2 = DateUtil.getJavaDate(OperandResolver.coerceValueToDouble(v_b));

                    switch( OperandResolver.coerceValueToString(args[2]).charAt(0)){
                        case 'm':
                            int m=  Months.monthsBetween( new DateTime(d1).withDayOfMonth(1), new DateTime(d2).withDayOfMonth(1)).getMonths();
                            return new NumberEval(m);
                        default:
                            System.out.println("DATEDIFF --> Unknown third parameter "+ OperandResolver.coerceValueToString(args[2]).charAt(0) );
                            return ErrorEval.NUM_ERROR;
                    }

                } catch (EvaluationException e) {
                    System.out.println("DateDiff Error --> "+args+ " "+ srcRowIndex+ "  "+srcColumnIndex);
                    e.printStackTrace();
                }
                return ErrorEval.NA;
            }
        });

        FunctionEval.registerFunction("DATEVALUE", new Function() {
            @Override
            public ValueEval evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
//                System.out.println("DATEVALUE Call --> "+ args + " "+ srcColumnIndex+" "+srcColumnIndex);

                ValueEval v = null;
                try {
                    v = OperandResolver.getSingleValue(args[0], srcRowIndex,srcColumnIndex);
                    System.out.println("args1 =" + OperandResolver.coerceValueToString(v));
                    if( v instanceof BlankEval)
                    {
                        System.out.println("DATEVALUE --> 0");
                        return new NumberEval(0);
                    }
                    else if ( v instanceof StringEval)
                    {
                        String vStr = OperandResolver.coerceValueToString(v);
                        if(vStr.length() != 0) {
                            DateFormat format = new SimpleDateFormat("yyyy-MM-dd");
                            Date date = format.parse(OperandResolver.coerceValueToString(v));
//                            System.out.println("DATEVALUE --> Date "+date.toString() );
                            return new NumberEval(DateUtil.getExcelDate(date));
                        }
                        else
                        {
                            System.out.println("DATEVALUE --> return num_error");
                            return ErrorEval.NUM_ERROR;
                        }
                    }
                } catch (EvaluationException e) {
                    e.printStackTrace();
                } catch (ParseException e) {
                    System.err.println(" value v "+v);
                    e.printStackTrace();
                }

//                System.out.println("DATEVALUE --> return NA");
                return ErrorEval.NA;
            }
        });



        //loop over code-tabellen , voor code-tabel <--> sheet-naam
        Sheet codeSheetReferentie = wb.getSheet("code-tabel-referentie");
        HashMap<String,String> cdtReferenties = new HashMap<String, String>();
        {
            int row=0;
            while(true){
                Row ro = codeSheetReferentie.getRow(row);
                if(ro == null)break;
                String sheetName = ro.getCell(2).getStringCellValue();
                String codetabelName = ro.getCell(1).getStringCellValue();
                cdtReferenties.put( sheetName, codetabelName);
                row++;
            }
        }

        HashMap<Integer,String> _t = new HashMap<Integer, String>();
        for( String shtName: cdtReferenties.keySet()){
            Sheet sheet = wb.getSheet(shtName);
            System.out.println("Sheet inlezen "+shtName);
            int row = 0;
            while(true){
                Row ro = sheet.getRow(row);
                if(ro==null)break;
                String key =  CellValueAsString(  ro.getCell(0));
                String val =  CellValueAsString(  ro.getCell(2));
                _t.put( Double.valueOf(Double.parseDouble(key)).intValue() ,val);
                row++;
            }
            codeTabellen.put( cdtReferenties.get(shtName), _t );
        }





//        {
//            Sheet sheet = wb.getSheet("deelberekening");
//
//            // perform debug output for the next evaluate-call only
//            evaluator.setDebugEvaluationOutputForNextEval(true);
//
//            CellReference cellReference = new CellReference("E17");
//            Row row = sheet.getRow(cellReference.getRow());
//            Cell cell = row.getCell(cellReference.getCol());
//
//            String v = getSingleCellValue( sheet, new pocRange1D("E17"));
//            System.out.println("val = "+v);
//            System.exit(0);
//        }


//            {
//
//
//            Sheet sheet = wb.getSheet("deelberekening");
//
//            // perform debug output for the next evaluate-call only
//            evaluator.setDebugEvaluationOutputForNextEval(true);
//
//            CellReference cellReference = new CellReference("E2");
//            Row row = sheet.getRow(cellReference.getRow());
//            Cell cell = row.getCell(cellReference.getCol());
//
//            if (cell!=null) {
//                switch (evaluator.evaluateInCell(cell).getCellType()) {
//                    case Cell.CELL_TYPE_BOOLEAN:
//                        System.out.println(cell.getBooleanCellValue());
//                        break;
//                    case Cell.CELL_TYPE_NUMERIC:
//                        System.out.println(cell.getNumericCellValue());
//                        break;
//                    case Cell.CELL_TYPE_STRING:
//                        System.out.println(cell.getStringCellValue());
//                        break;
//                    case Cell.CELL_TYPE_BLANK:
//                        break;
//                    case Cell.CELL_TYPE_ERROR:
//                        System.out.println(cell.getErrorCellValue());
//                        break;
//
//                    // CELL_TYPE_FORMULA will never occur
//                    case Cell.CELL_TYPE_FORMULA:
//                        break;
//                }
//            }
//
//            evaluator.evaluateFormulaCell(cell);
//            evaluator.evaluateFormulaCell(cell);
//        }


        if(sheet_validator(wb)==false) {
            System.out.println("Sheet error");
            System.exit(1);
        }
        else {
            System.out.println("Sheet OK");
        }

        Sheet extractieSheet = wb.getSheet("extractie-dln");
        Sheet extractieNprOpRenteSheet = wb.getSheet("extractie-npr-op-rente");
        Sheet extractieConfrontatie = wb.getSheet("Opzet_confrontatie");




        CellRangeAddress extractieRangeNAW =  CellRangeAddress.valueOf("F3:F19");
        CellRangeAddress extractieRangeDVD =  CellRangeAddress.valueOf("N25:W200");

        resetRange( extractieSheet,extractieRangeNAW );
        resetRange( extractieSheet,extractieRangeDVD );

        System.out.println("Uitvoer "+ getSingleCellValue( extractieConfrontatie, new pocRange1D("E8")) );


    }

    private static boolean sheet_validator(Workbook wb){

        boolean b=true;
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        for (Sheet sheet : wb) {
            for (Row r : sheet) {
                for (Cell c : r) {
                    if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {

                        try {
                            evaluator.evaluateFormulaCell(c);
                        }
                        catch( FormulaParseException e){
                            System.out.println("Sheet "+sheet.getSheetName()+" Cell under eval "+c.getAddress().toString());
                            System.out.println("--> "+ e.toString());
                            b=false;
                        }
                    }
                }
            }
        }

        return b;
    }

    private static void resetRange( Sheet sht,CellRangeAddress range){

        Cell clr = null;
        Row  row = null;
        for ( int r= range.getFirstRow(); r<=range.getLastRow();r++ ){
            row = sht.getRow(r);
            for( int c=range.getFirstColumn(); c<=range.getLastColumn(); c++){
               clr = row.getCell(c, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if( clr != null){
                    evaluator.notifyUpdateCell(clr);
                }
            }
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


    public static String getSingleCellValue(Sheet sht, pocRange1D range){

        Row r = sht.getRow( range.row );
        String val = CellValueAsString( r.getCell( range.column));

        return val;
    }

    private static Cell getSingleCell(Sheet sht, pocRange1D range){
        Row r = sht.getRow( range.row);

        if( r==null){
            r=sht.createRow(range.row);
        }

        Cell c = r.getCell( range.column);
        return c;
    }


    public static Map<String, String> getInvoer(Sheet invoer_sht, int invoerNr ){

        HashMap hm = new HashMap<String,String>();

        Row invoerRow=null;


        // Zoek juiste row bij invoerNr. Invoer_BPFP nr zit in col A (=0)

        for(int i=2;i<100;i++){
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
            if( valString != null && valString.length() > 0){
                String tagString = CellValueAsString( tagrow.getCell(col));
                if(hm.containsKey(tagString)){
                    hm.put(tagString+"_"+col, valString);
                }
                else {
                    hm.put(tagString, valString);
                }
            }
        }

        return hm;
    }

}

