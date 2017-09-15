package nl.tkp.pps.sds;

import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.functions.FreeRefFunction;


public class udf_DATEDIF implements FreeRefFunction {


    @Override
    public ValueEval evaluate( ValueEval[] args, OperationEvaluationContext ec ) {
        if (args.length != 3) {
            return ErrorEval.VALUE_INVALID;
        }

        //org.apache.poi.ss.formula.LazyAreaEval le = ( org.apache.poi.ss.formula.LazyAreaEval )args[1];

        // Arg 1 = Codetabel
        // Arg 2 = index
        int  result=-99 ;

//        NumberEval kolomnummer;
//        LazyRefEval leeftijd;
//        ValueEval tabel;

//        try {
//            if (args[2] instanceof NumberEval) {
//                kolomnummer = (NumberEval) args[2];
//                //System.out.println("kolomnummer =" + kolomnummer.getNumberValue());
//            } else {
//                return ErrorEval.NUM_ERROR;
//            }


//            Vlookup vl = new Vlookup();
//            ValueEval tar1 = vl.evaluate( ec.getRowIndex(), ec.getColumnIndex(), new NumberEval(x1),tabel,kolomnummer, BoolEval.FALSE);
//            ValueEval tar2 = vl.evaluate( ec.getRowIndex(), ec.getColumnIndex(), new NumberEval(x2),tabel,kolomnummer, BoolEval.FALSE);

//            double tar1_double = OperandResolver.coerceValueToDouble(tar1);
//


//            result=tarief;
            //checkValue(result);

//        } catch (EvaluationException e) {
//            e.printStackTrace() ;
//            return e.getErrorEval();
//        }

        return new NumberEval( result ) ;
    }


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
