/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ars.traco.databeans.n414;

/**
 *
 * @author Sreedarshs
 */
public class L2 {
    private int matches;
    
    private OvertredingenType overtredingenType=new OvertredingenType();
    private PerformanceType performanceType=new PerformanceType();
    private SnelhedenType snelheden=new SnelhedenType();
    private PassagesType passagesType=new PassagesType();

    public L2() {
    }

    public L2( int matches) {
        this();
        this.matches = matches;
    }

    public int getMatches() {
        return matches;
    }

    public void setMatches(int matches) {
        this.matches = matches;
    }

    public OvertredingenType getOvertredingenType() {
        return overtredingenType;
    }

    public void setOvertredingenType(OvertredingenType overtredingenType) {
        this.overtredingenType = overtredingenType;
    }

    public PerformanceType getPerformanceType() {
        return performanceType;
    }

    public void setPerformanceType(PerformanceType performanceType) {
        this.performanceType = performanceType;
    }

    public SnelhedenType getSnelheden() {
        return snelheden;
    }

    public void setSnelheden(SnelhedenType snelheden) {
        this.snelheden = snelheden;
    }

    public PassagesType getpassagesType() {
        return passagesType;
    }

    public void setpassagesType(PassagesType passagesType) {
        this.passagesType = passagesType;
    }
    
    @Override
    public String toString() {
        return "Section{" +  ", matches=" + matches + ", overtredingenType=" + overtredingenType + ", performanceType=" + performanceType + ", snelheden=" + snelheden + '}';
    }
}