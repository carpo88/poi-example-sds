package nl.tkp.pps.sds;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class pocRange1D {

    static Pattern rngRE = Pattern.compile("^([A-Z]+)(\\d+)");

    int row=-1;
    int column=-1;

    private static int colIndex(String col)
    {   int index=0;
        int mul=0;
        for(int i=col.length()-1;i>=0;i--)
        {
            index  += (col.charAt(i)-64) * Math.pow(26, mul);
            mul++;
        }
        return index;
    }

    public pocRange1D(String rngStr)    {

        Matcher rngM = rngRE.matcher(rngStr);
        if(rngM.find()){
            row += Integer.parseInt( rngM.group(2));
            column += colIndex( rngM.group(1));
        }


    }

}
