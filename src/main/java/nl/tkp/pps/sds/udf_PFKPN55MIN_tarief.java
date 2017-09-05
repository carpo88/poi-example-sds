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


    @Override
    public ValueEval evaluate( ValueEval[] args, OperationEvaluationContext ec ) {
        if (args.length != 3) {
            return ErrorEval.VALUE_INVALID;
        }

        //org.apache.poi.ss.formula.LazyAreaEval le = ( org.apache.poi.ss.formula.LazyAreaEval )args[1];

        // Arg 1 = KPN55Min:C11  -> leeftijd
        // Arg 2 = Factoren!A5:AO56 = range -> tabel
        // Arg 3 = NumberEval[5] --> kolomnr ??
        int leeftijd;
        double   tabel, kolomnummer,result ;

        //double n  = 1 - VindRest(tabel, leeftijd) --> VLookup(leeftijd, tabel, 1)



        result=-1234.567;
        try {
            if(args[0] instanceof org.apache.poi.ss.formula.LazyRefEval){
                LazyRefEval re = (LazyRefEval)args[0];
                ValueEval v1 = re.getInnerValueEval( ec.getSheetIndex() );
                leeftijd = OperandResolver.coerceValueToInt( v1);
                System.out.println("Leeftijd ="+leeftijd);

            }

            if (args[1] instanceof AreaEval) {
                AreaEval ae = (AreaEval) args[1];
                for(int i= ae.getFirstRow();i!=ae.getLastRow();i++)
                {
                    ValueEval v2 = ae.getRelativeValue(i,0);
                    int  indexValue = OperandResolver.coerceValueToInt(v2);
                    if ( indexValue == leeftijd){
                        ValueEval v3 = ae.getRelativeValue(i,)
                    }
                }
            }

            ValueEval v1 = OperandResolver.getSingleValue( args[0],
                    ec.getRowIndex(),
                    ec.getColumnIndex() ) ;
            ValueEval v2 = OperandResolver.getSingleValue( args[1],
                    ec.getRowIndex(),
                    ec.getColumnIndex() ) ;
            ValueEval v3 = OperandResolver.getSingleValue( args[2],
                    ec.getRowIndex(),
                    ec.getColumnIndex() ) ;

//            principal  = OperandResolver.coerceValueToDouble( v1 ) ;
//            rate  = OperandResolver.coerceValueToDouble( v2 ) ;
//            years = OperandResolver.coerceValueToDouble( v3 ) ;
//
//            result = calculateMortgagePayment( principal, rate, years ) ;

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
