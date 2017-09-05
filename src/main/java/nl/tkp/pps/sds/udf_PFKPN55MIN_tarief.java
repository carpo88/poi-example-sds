package nl.tkp.pps.sds;

import org.apache.poi.ss.formula.LazyRefEval;
import org.apache.poi.ss.formula.OperationEvaluationContext ;
import org.apache.poi.ss.formula.TwoDEval;
import org.apache.poi.ss.formula.eval.*;
import org.apache.poi.ss.formula.functions.FreeRefFunction ;
import org.apache.poi.ss.formula.functions.Vlookup;
import org.apache.poi.ss.usermodel.CellRange;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;


public class udf_PFKPN55MIN_tarief implements FreeRefFunction {

    /*
            hulp = Application.VLookup(leeftijd, tabel, 1)

            hulp = hulp - Application.RoundDown(hulp, 0)
            VindRest = hulp
     */
    private double VindRest(OperationEvaluationContext ec,ValueEval tabel, LazyRefEval leeftijd){

        double hulp=0;

        //System.out.println("DBG= row="+ec.getRowIndex() +" col="+ec.getColumnIndex()+ " sht="+ec.getSheetIndex()+" tbl="+tabel.toString() + " lft="+leeftijd.toString() );
        Vlookup vl = new Vlookup();
        // ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval lookup_value, ValueEval table_array,
        //                    ValueEval col_index, ValueEval range_lookup)
        ValueEval hulp_ve = vl.evaluate( ec.getRowIndex(), ec.getColumnIndex(), leeftijd, tabel, new NumberEval(1), BoolEval.TRUE);

        try {
            if(hulp_ve instanceof org.apache.poi.ss.formula.eval.ErrorEval){
                System.out.println(" error value ");
            }
            hulp = OperandResolver.coerceValueToDouble(hulp_ve);
            hulp = hulp - Math.abs(hulp);

        } catch (EvaluationException e) {
            e.printStackTrace();
        }

        return hulp;
    }

    @Override
    public ValueEval evaluate( ValueEval[] args, OperationEvaluationContext ec ) {
        if (args.length != 3) {
            return ErrorEval.VALUE_INVALID;
        }

        //org.apache.poi.ss.formula.LazyAreaEval le = ( org.apache.poi.ss.formula.LazyAreaEval )args[1];

        // Arg 1 = KPN55Min:C11  -> leeftijd
        // Arg 2 = Factoren!A5:AO56 = range -> tabel
        // Arg 3 = NumberEval[5] --> kolomnr ??
        double  result ;

        NumberEval kolomnummer;
        LazyRefEval leeftijd;
        ValueEval tabel;

        try {
            if (args[2] instanceof org.apache.poi.ss.formula.eval.NumberEval) {
                kolomnummer = (NumberEval) args[2];
                //System.out.println("kolomnummer =" + kolomnummer.getNumberValue());
            } else {
                return ErrorEval.NUM_ERROR;
            }

            if(args[0] instanceof org.apache.poi.ss.formula.LazyRefEval) {
                leeftijd = (LazyRefEval) args[0];
                //System.out.println("Leeftijd =" + leeftijd);
            } else {
                return ErrorEval.NUM_ERROR;
            }

            //if(args[1] instanceof org.apache.poi.ss.formula.LazyAreaEval) {
                tabel = (ValueEval) args[1];
                //System.out.println("tabel =" +tabel.toString());
            //} else {
            //    return ErrorEval.NUM_ERROR;
            //}

            double n = 1 - VindRest(ec,tabel,leeftijd);
            ValueEval leeftijdValue  = OperandResolver.getSingleValue(leeftijd, ec.getRowIndex(), ec.getRowIndex());
            double leeftijd_d = OperandResolver.coerceValueToDouble(leeftijdValue);

            double x1 = Math.floor(  leeftijd_d + n) -n;
            double x2 = Math.ceil( leeftijd_d + n) -n;

            double r = leeftijd_d - x1;

            Vlookup vl = new Vlookup();
            ValueEval tar1 = vl.evaluate( ec.getRowIndex(), ec.getColumnIndex(), new NumberEval(x1),tabel,kolomnummer, BoolEval.FALSE);
            ValueEval tar2 = vl.evaluate( ec.getRowIndex(), ec.getColumnIndex(), new NumberEval(x2),tabel,kolomnummer, BoolEval.FALSE);
            double tar1_double = OperandResolver.coerceValueToDouble(tar1);
            double tar2_double = OperandResolver.coerceValueToDouble(tar2);

            double tarief = ( 1 - r) * tar1_double + r * tar2_double;

        /*

        n = 1 - VindRest(tabel, leeftijd)



        x1 = Application.RoundDown(leeftijd + n, 0) - n
        x2 = Application.RoundUp(leeftijd + n, 0) - n
        r = leeftijd - x1
        tar1 = Application.VLookup(x1, tabel, kolomnr, False)
        tar2 = Application.VLookup(x2, tabel, kolomnr, False)
        tarief = (1 - r) * tar1 + r * tar2

       */

            result=tarief;
            checkValue(result);

        } catch (EvaluationException e) {
            e.printStackTrace() ;
            return e.getErrorEval();
        }

        return new NumberEval( result ) ;
    }

//    public double calculateMortgagePayment( double p, double r, double y ) {
//        double i = r / 12 ;
//        double n = y * 12 ;
//
//        //M = P [ i(1 + i)n ] / [ (1 + i)n - 1]
//        double principalAndInterest =
//                p * (( i * Math.pow((1 + i),n ) ) / ( Math.pow((1 + i),n) - 1))  ;
//
//        return principalAndInterest ;
//    }

    /**
     * Excel does not support infinities and NaNs, rather, it gives a #NUM! error in these cases
     *
     * @throws EvaluationException (#NUM!) if <tt>result</tt> is <tt>NaN</> or <tt>Infinity</tt>
     */
    static final void checkValue(double result) throws EvaluationException {
        if (Double.isNaN(result) || Double.isInfinite(result)) {
            throw new EvaluationException(ErrorEval.NUM_ERROR);
        }
    }

}
