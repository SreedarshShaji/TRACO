/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ars.test;

import java.text.DecimalFormat;

/**
 *
 * @author ADMIN
 */
public class Test1 {
    public static void main(String[] args) {
        double no=12.7;
        DecimalFormat dec = new DecimalFormat("#0.00");
        String output=dec.format(no);
        System.out.println(output);
    }
}
