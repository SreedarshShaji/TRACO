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


public class PerformanceType{


    private String tijdvolledigbeschikbaar;
    private double overtredingenratio;
    private double autoratio;
    private double handhaafratio;
    private double matchratio;

    public PerformanceType() {
        super();
    }

    //Setters and Getters

    public String getTijdvolledigbeschikbaar() {
        return tijdvolledigbeschikbaar;
    }

    public void setTijdvolledigbeschikbaar(String tijdvolledigbeschikbaar) {
        this.tijdvolledigbeschikbaar = tijdvolledigbeschikbaar;
    }

    public double getOvertredingenratio() {
        return overtredingenratio;
    }

    public void setOvertredingenratio(double overtredingenratio) {
        this.overtredingenratio = overtredingenratio;
    }

    public double getAutoratio() {
        return autoratio;
    }

    public void setAutoratio(double autoratio) {
        this.autoratio = autoratio;
    }

    public double getHandhaafratio() {
        return handhaafratio;
    }

    public void setHandhaafratio(double handhaafratio) {
        this.handhaafratio = handhaafratio;
    }

    public double getMatchratio() {
        return matchratio;
    }

    public void setMatchratio(double matchratio) {
        this.matchratio = matchratio;
    }
    
    //to string method

    @Override
    public String toString() {
        return "N414{" + "tijdvolledigbeschikbaar=" + tijdvolledigbeschikbaar + ", overtredingenratio=" + overtredingenratio + ", autoratio=" + autoratio + ", handhaafratio=" + handhaafratio + ", matchratio=" + matchratio + '}';
    }
    
}
