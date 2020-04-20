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
public class OvertredingenType {

    private int overtredingenTotaal;
    private int auto;
    private int dubbeleOvertredingenPardon;
    private int overigPardon;
    private int hand;

    public OvertredingenType() {
    }

    public OvertredingenType(int overtredingenTotaal, int auto, int dubbeleOvertredingenPardon, int overigPardon, int hand) {
        this.overtredingenTotaal = overtredingenTotaal;
        this.auto = auto;
        this.dubbeleOvertredingenPardon = dubbeleOvertredingenPardon;
        this.overigPardon = overigPardon;
        this.hand = hand;
    }

    /* Getters and setters */
    public int getOvertredingenTotaal() {
        return overtredingenTotaal;
    }

    public void setOvertredingenTotaal(int overtredingenTotaal) {
        this.overtredingenTotaal = overtredingenTotaal;
    }

    public int getAuto() {
        return auto;
    }

    public void setAuto(int auto) {
        this.auto = auto;
    }

    public int getDubbeleOvertredingenPardon() {
        return dubbeleOvertredingenPardon;
    }

    public void setDubbeleOvertredingenPardon(int dubbeleOvertredingenPardon) {
        this.dubbeleOvertredingenPardon = dubbeleOvertredingenPardon;
    }

    public int getOverigPardon() {
        return overigPardon;
    }

    public void setOverigPardon(int overigPardon) {
        this.overigPardon = overigPardon;
    }

    public int getHand() {
        return hand;
    }

    public void setHand(int hand) {
        this.hand = hand;
    }
    /* To Stting Method*/
    @Override
    public String toString() {
        return "OvertredingenType{" + "overtredingenTotaal=" + overtredingenTotaal + ", auto=" + auto + ", dubbeleOvertredingenPardon=" + dubbeleOvertredingenPardon + ", overigPardon=" + overigPardon + ", hand=" + hand + '}';
    }

}
