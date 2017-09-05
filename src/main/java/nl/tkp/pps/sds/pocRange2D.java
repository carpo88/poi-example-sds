package nl.tkp.pps.sds;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class pocRange2D{

    static Pattern rngRE = Pattern.compile("^([A-Z]+)(\\d+):([A-Z]+)(\\d+)");

    int startRow=-1;
    int endRow=-1;
    int startColumn=-1;
    int endColumn=-1;

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

    public  pocRange2D ( String rngStr)    {

        Matcher rngM = rngRE.matcher(rngStr);
        if(rngM.find()){
            startRow += Integer.parseInt( rngM.group(2));
            startColumn += colIndex( rngM.group(1));
            endRow += Integer.parseInt( rngM.group(4));
            endColumn += colIndex(rngM.group(3));
        }


    }

}
