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
public class PassagesType {
    
    private int totalUit=0;
    private int totalIn=0;

    public PassagesType() {
    }    
    
    public int getTotalUit() {
        return totalUit;
    }

    public void setTotalUit(int totalUit) {
        this.totalUit = totalUit;
    }

    public int getTotalIn() {
        return totalIn;
    }

    public void setTotalIn(int totalIn) {
        this.totalIn = totalIn;
    }

    @Override
    public String toString() {
        return "Passages{" + "totalUit=" + totalUit + ", totalIn=" + totalIn + '}';
    }
}