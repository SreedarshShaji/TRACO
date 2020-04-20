/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ars.traco.filehandler;

/**
 *
 * @author Sreedarshs
 */
public class fileChecker {
    
    public boolean checkXMLName(String fileName)
    {
        boolean result=false;
        
        if(fileName.indexOf("TRACO")<0)
        {
            result=false;
        }
        else if(fileName.indexOf("Daily_Report")<0)
        {
            result=false;
        }
        else
        {
            result=true;
        }
        
        return result;
    }
    
}
