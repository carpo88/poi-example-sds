package nl.tkp.pps.sds;

import org.joda.time.DateTime;

public class Invoer_BPFP {
    public final int invoernr;
    public final int pensioenNummer;
    public final String geslacht;
    public final DateTime geboorteDatum;
    public final DateTime rekenDatum;
    public final DateTime betaalDatum;
    public final double uurLoon;
    public final double gewerkteUrenEersteVolledigePeriode;
    public final double normUrenPerWeek;
    public final String verloningsPeriode;
    public final String  maandEersteLoongegevens;
    public final double overdrachtsWaardeBPP;
    public final double  overdrachtsWaardeOP;
    public final double overdrachtsWaardeVerevendOP;
    public final int werkgeversNummer;
    public final String  circuit;
    public final double pensioenLeeftijd;

    private Invoer_BPFP(InvoerBuilder builder){
        this.invoernr = builder.invoernr;
        this.pensioenNummer = builder.pensioenNummer;
        this.geslacht = builder.geslacht;
        this.geboorteDatum = builder.geboorteDatum;
        this.rekenDatum = builder.rekenDatum;
        this.betaalDatum = builder.betaalDatum;
        this.uurLoon = builder.uurLoon;
        this.gewerkteUrenEersteVolledigePeriode = builder.gewerkteUrenEersteVolledigePeriode;
        this.normUrenPerWeek = builder.normUrenPerWeek;
        this.verloningsPeriode = builder.verloningsPeriode;
        this.maandEersteLoongegevens = builder.maandEersteLoongegevens;
        this.overdrachtsWaardeBPP = builder.overdrachtsWaardeBPP;
        this.overdrachtsWaardeOP = builder.overdrachtsWaardeOP;
        this.overdrachtsWaardeVerevendOP = builder.overdrachtsWaardeVerevendOP;
        this.werkgeversNummer = builder.werkgeversNummer;
        this.circuit = builder.circuit;
        this.pensioenLeeftijd = builder.pensioenLeeftijd;

    }


    public static class InvoerBuilder {
        private  int invoernr;
        private final int pensioenNummer;
        private final String geslacht;
        private final DateTime geboorteDatum;
        private  DateTime rekenDatum;
        private  DateTime betaalDatum;
        private  double uurLoon;
        private  double gewerkteUrenEersteVolledigePeriode;
        private  double normUrenPerWeek;
        private  String verloningsPeriode;
        private  String  maandEersteLoongegevens;
        private  double overdrachtsWaardeBPP;
        private  double  overdrachtsWaardeOP;
        private double overdrachtsWaardeVerevendOP;
        private  int werkgeversNummer;
        private  String  circuit;
        private  double pensioenLeeftijd;


        public InvoerBuilder(int pensioenNummer, String geslacht, DateTime geboorteDatum){
            this.pensioenNummer = pensioenNummer;
            this.geslacht = geslacht;
            this.geboorteDatum = geboorteDatum;
        }

        public InvoerBuilder rekenDatum(DateTime rekenDatum){
            this.rekenDatum = rekenDatum;
            return this;
        }

        public InvoerBuilder betaalDatum(DateTime betaalDatum){
            this.betaalDatum = betaalDatum;
            return this;
        }

        public InvoerBuilder uurLoon (double uurLoon){
            this.uurLoon = uurLoon;
            return this;
        }

        public InvoerBuilder gewerkteUrenEersteVolledigePeriode (double gewerkteUrenEersteVolledigePeriode){
            this.gewerkteUrenEersteVolledigePeriode = gewerkteUrenEersteVolledigePeriode;
            return this;
        }

        public InvoerBuilder normUrenPerWeek (double normUrenPerWeek){
            this.normUrenPerWeek = normUrenPerWeek;
            return this;
        }

        public InvoerBuilder verloningsPeriode (String verloningsPeriode){
            this.verloningsPeriode = verloningsPeriode;
            return this;
        }

        public InvoerBuilder maandEersteLoongegevens (String maandEersteLoongegevens){
            this.maandEersteLoongegevens = maandEersteLoongegevens;
            return this;
        }

        public InvoerBuilder overdrachtsWaardeBPP (double overdrachtsWaardeBPP){
            this.overdrachtsWaardeBPP = overdrachtsWaardeBPP;
            return this;
        }

        public InvoerBuilder overdrachtsWaardeOP (double overdrachtsWaardeOP){
            this.overdrachtsWaardeOP = overdrachtsWaardeOP;
            return this;
        }

        public InvoerBuilder overdrachtsWaardeVerevendOP (double overdrachtsWaardeVerevendOP){
            this.overdrachtsWaardeVerevendOP = overdrachtsWaardeVerevendOP;
            return this;
        }

        public InvoerBuilder werkgeversNummer (int werkgeversNummer){
            this.werkgeversNummer = werkgeversNummer;
            return this;
        }

        public InvoerBuilder circuit (String circuit){
            this.circuit = circuit;
            return this;
        }

        public InvoerBuilder pensioenLeeftijd (double pensioenLeeftijd){
            this.pensioenLeeftijd = pensioenLeeftijd;
            return this;
        }


        public Invoer_BPFP build(){
            return new Invoer_BPFP(this);
        }

    }

}
