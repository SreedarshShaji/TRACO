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
public class FileClassifier {
    
    public String fileClassifier(String fileAbsolutePath)
    {
        String result="";
        String fileName=fileAbsolutePath.substring(fileAbsolutePath.lastIndexOf("\\")+1);
        System.out.println(fileName);
        
        if(!(fileName.indexOf("N414")<0))
        {
            result="N414";
        }
        if(!(fileName.indexOf("N333")<0))
        {
            result="N333";
        }
        
        return result;
    }
    
}