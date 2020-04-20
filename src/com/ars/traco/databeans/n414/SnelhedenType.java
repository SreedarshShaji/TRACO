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
public class SnelhedenType {
    private double gemiddeld;
    private double max;

    public SnelhedenType() {
    }

    public SnelhedenType(double gemiddeld, double max) {
        this.gemiddeld = gemiddeld;
        this.max = max;
    }

    /* Getters and setters */
    public double getGemiddeld() {
        return gemiddeld;
    }

    public void setGemiddeld(double gemiddeld) {
        this.gemiddeld = gemiddeld;
    }

    public double getMax() {
        return max;
    }

    public void setMax(double max) {
        this.max = max;
    }

    @Override
    public String toString() {
        return "Snelheden{" + "gemiddeld=" + gemiddeld + ", max=" + max + '}';
    }
    
    
}
