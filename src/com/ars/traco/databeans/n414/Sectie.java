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
public class Sectie {
    R1 r1=new R1();
    L2 l2=new L2();

    String date;

    public Sectie() {
    }

    public Sectie(String date) {
        this.date = date;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }
    
    
    public R1 getR1() {
        return r1;
    }

    public void setR1(R1 r1) {
        this.r1 = r1;
    }

    public L2 getL2() {
        return l2;
    }

    public void setL2(L2 l2) {
        this.l2 = l2;
    }

    @Override
    public String toString() {
        return "Sectie{" + "r1=" + r1 + ", l2=" + l2 + ", date=" + date + '}';
    }
    
    
}
