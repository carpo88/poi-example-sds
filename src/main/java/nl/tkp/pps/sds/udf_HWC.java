package nl.tkp.pps.sds;

import org.apache.poi.ss.formula.DataValidationEvaluator;
import org.apache.poi.ss.formula.LazyRefEval;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.*;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.functions.Vlookup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;


public class udf_HWC implements FreeRefFunction {


    @Override
    public ValueEval evaluate( ValueEval[] args, OperationEvaluationContext ec ) {
        if (args.length != 2) {
            return ErrorEval.VALUE_INVALID;
        }

        String codetabel=null;
        String index=null;
        //org.apache.poi.ss.formula.LazyAreaEval le = ( org.apache.poi.ss.formula.LazyAreaEval )args[1];

        // Arg 1 = Codetabel
        // Arg 2 = index
        double  result;

//        NumberEval kolomnummer;
//        LazyRefEval leeftijd;
//        ValueEval tabel;

        try {
            if (args[0] instanceof StringEval) {
                StringEval cdtEval = (StringEval) args[0];
                codetabel = cdtEval.getStringValue();
                System.out.println("args0 = "+codetabel);
            } else
            if(args[0] instanceof org.apache.poi.ss.formula.LazyRefEval) {
                LazyRefEval cdtEval = (LazyRefEval) args[0];
                ValueEval v = OperandResolver.getSingleValue(cdtEval, ec.getRowIndex(),ec.getColumnIndex());
                System.out.println("args0 =" + OperandResolver.coerceValueToString(v));
                codetabel = OperandResolver.coerceValueToString(v);
            } else {
                return ErrorEval.NUM_ERROR;
            }


            if (args[1] instanceof StringEval) {
                StringEval cdtEval = (StringEval) args[0];
                index = cdtEval.getStringValue();
            } else
            if(args[1] instanceof org.apache.poi.ss.formula.LazyRefEval) {
                LazyRefEval cdtEval = (LazyRefEval) args[1];
                ValueEval v = OperandResolver.getSingleValue(cdtEval, ec.getRowIndex(),ec.getColumnIndex());
                System.out.println("args1 =" + OperandResolver.coerceValueToString(v));
                index = OperandResolver.coerceValueToString(v);
            } else {
                return ErrorEval.NUM_ERROR;
            }

            HashMap<Integer,String> c = pocPOIKrokus.codeTabellen.get(codetabel);
            if(c==null)return ErrorEval.CIRCULAR_REF_ERROR;

            String s = c.get( Integer.parseInt(index) );
            if(s==null)return ErrorEval.NAME_INVALID;
            result = Double.parseDouble( s);

            checkValue(result);

        } catch (EvaluationException e) {
            e.printStackTrace() ;
            return e.getErrorEval();
        }

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
