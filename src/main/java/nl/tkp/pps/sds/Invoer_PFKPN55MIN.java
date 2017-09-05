package nl.tkp.pps.sds;

import org.apache.poi.poifs.filesystem.OPOIFSDocument;
import org.joda.time.DateTime;

public class Invoer_PFKPN55MIN {
    public final int invoernr;
    public final String geslacht;
    public final String burgerlijkeStaat;
    public final DateTime geboorteDatum;
    public final DateTime pensioenDatum;
    public final DateTime uitdienstDatum;
    public final DateTime rekenDatum;
    public final DateTime betaalDatum;
    public final String regeling;
    public final String depot;

    public final double LOP67_excl_verevening;
    public final double verevend_LOP67;
    public final double OPO;
    public final double PP_middelloon_excl_BPP;
    public final double PP_bijzonder_LOP;
    public final double PP_eindloon_oud_excl_BPP;
    public final double PP_bijzonder_oud;
    public final double BPR;
    public final double IPS;
    public final double PPS;

    private Invoer_PFKPN55MIN(InvoerBuilder builder){
        this.invoernr = builder.invoernr;

        this.geslacht = builder.geslacht;
        this.burgerlijkeStaat = builder.burgerlijkeStaat;
        this.geboorteDatum = builder.geboorteDatum;
        this.pensioenDatum = builder.pensioenDatum;
        this.uitdienstDatum = builder.uitdienstDatum;
        this.rekenDatum = builder.rekenDatum;
        this.betaalDatum = builder.betaalDatum;
        this.regeling = builder.regeling;
        this.depot = builder.depot;
        this.LOP67_excl_verevening = builder.LOP67_excl_verevening;
        this.verevend_LOP67 = builder.verevend_LOP67;
        this.OPO = builder.OPO;
        this.PP_middelloon_excl_BPP = builder.PP_middelloon_excl_BPP;
        this.PP_bijzonder_LOP = builder.PP_bijzonder_LOP;
        this.PP_eindloon_oud_excl_BPP = builder.PP_eindloon_oud_excl_BPP;
        this.PP_bijzonder_oud = builder.PP_bijzonder_oud;
        this.BPR = builder.BPR;
        this.IPS = builder.IPS;
        this.PPS = builder.PPS;


    }


    public static class InvoerBuilder {
        private int invoernr;
        private String geslacht;
        private String burgerlijkeStaat;
        private DateTime geboorteDatum;
        private DateTime pensioenDatum;
        private DateTime uitdienstDatum;
        private DateTime rekenDatum;
        private DateTime betaalDatum;
        private String regeling;
        private String depot;
        private double LOP67_excl_verevening;
        private double verevend_LOP67;
        private double OPO;
        private double PP_middelloon_excl_BPP;
        private double PP_bijzonder_LOP;
        private double PP_eindloon_oud_excl_BPP;
        private double PP_bijzonder_oud;
        private double BPR;
        private double IPS;
        private double PPS;

        public InvoerBuilder invoernr (int invoernr){
            this.invoernr = invoernr;
            return this;
        }

        public InvoerBuilder geslacht (String geslacht){
            this.geslacht = geslacht;
            return this;
        }

        public InvoerBuilder burgerlijkeStaat (String burgerlijkeStaat){
            this.burgerlijkeStaat = burgerlijkeStaat;
            return this;
        }

        public InvoerBuilder geboorteDatum (DateTime geboorteDatum){
            this.geboorteDatum = geboorteDatum;
            return this;
        }

        public InvoerBuilder pensioenDatum (DateTime pensioenDatum){
            this.pensioenDatum = pensioenDatum;
            return this;
        }

        public InvoerBuilder uitdienstDatum (DateTime uitdienstDatum){
            this.uitdienstDatum = uitdienstDatum;
            return this;
        }

        public InvoerBuilder rekenDatum (DateTime rekenDatum){
            this.rekenDatum = rekenDatum;
            return this;
        }

        public InvoerBuilder betaalDatum (DateTime betaalDatum){
            this.betaalDatum = betaalDatum;
            return this;
        }

        public InvoerBuilder regeling (String regeling){
            this.regeling = regeling;
            return this;
        }

        public InvoerBuilder depot (String depot){
            this.depot = depot;
            return this;
        }

        public InvoerBuilder LOP67_excl_verevening (double LOP67_excl_verevening){
            this.LOP67_excl_verevening = LOP67_excl_verevening;
            return this;
        }

        public InvoerBuilder verevend_LOP67 (double verevend_LOP67){
            this.verevend_LOP67 = verevend_LOP67;
            return this;
        }

        public InvoerBuilder OPO (double OPO){
            this.OPO = OPO;
            return this;
        }

        public InvoerBuilder PP_middelloon_excl_BPP (double PP_middelloon_excl_BPP){
            this.PP_middelloon_excl_BPP = PP_middelloon_excl_BPP;
            return this;
        }

        public InvoerBuilder PP_eindloon_oud_excl_BPP (double PP_eindloon_oud_excl_BPP){
            this.PP_eindloon_oud_excl_BPP = PP_eindloon_oud_excl_BPP;
            return this;
        }

        public InvoerBuilder PP_bijzonder_LOP (double PP_bijzonder_LOP){
            this.PP_bijzonder_LOP = PP_bijzonder_LOP;
            return this;
        }

        public InvoerBuilder PP_bijzonder_oud (double PP_bijzonder_oud){
            this.PP_bijzonder_oud = PP_bijzonder_oud;
            return this;
        }


        public InvoerBuilder BPR (double BPR){
            this.BPR = IPS;
            return this;
        }


        public InvoerBuilder IPS (double IPS){
            this.IPS = IPS;
            return this;
        }

        public InvoerBuilder PPS (double PPS){
            this.PPS = PPS;
            return this;
        }


        public Invoer_PFKPN55MIN build(){
            return new Invoer_PFKPN55MIN(this);
        }

    }

}
