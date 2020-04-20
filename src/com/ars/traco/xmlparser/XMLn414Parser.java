/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ars.traco.xmlparser;

import com.ars.traco.databeans.n414.L2;
import com.ars.traco.databeans.n414.OvertredingenType;
import com.ars.traco.databeans.n414.PassagesType;
import com.ars.traco.databeans.n414.PerformanceType;
import com.ars.traco.databeans.n414.R1;
import com.ars.traco.databeans.n414.Sectie;
import com.ars.traco.databeans.n414.SnelhedenType;
import java.io.FileInputStream;
import com.ars.json.JSONArray;
import com.ars.json.JSONObject;
import com.ars.json.XML;


/*
 * @author Sreedarshs
 */
public class XMLn414Parser {

    public static int INDENTATION = 4;
    public Sectie mainSectie=null;


	public XMLn414Parser(String filePath) {

        try {

            /*Reads the file content from the file path*/
            FileInputStream fis = new FileInputStream(filePath);
            byte[] buffer = new byte[10];
            StringBuilder sb = new StringBuilder();
            while (fis.read(buffer) != -1) {
                sb.append(new String(buffer));
                buffer = new byte[10];
            }
            fis.close();

            /* Convert the file content to JSON */
            String fileContent = sb.toString();
            JSONObject jsonObj = XML.toJSONObject(fileContent);
            String json = jsonObj.toString(INDENTATION);

            /* Extract json file */
            JSONObject obj = new JSONObject(json);
            String date = obj.getJSONObject("TCS-dagrapport").getString("datum");

            /* Objects for holding R1 */
            R1 r1 = new R1();

            PerformanceType R1PerformanceType = new PerformanceType();
            OvertredingenType R1OvertredingenType = new OvertredingenType();
            SnelhedenType R1SnelhedenType = new SnelhedenType();
            PassagesType R1PassagesType = new PassagesType();

            /* Reading R - 1 Details */
            JSONArray sectie = obj.getJSONObject("TCS-dagrapport").getJSONArray("sectie");
            JSONObject r1SectieStatistiek = sectie.getJSONObject(0).getJSONObject("Sectie-statistiek");

            /* Extracting JSON subsections */
            JSONObject r1Performance = r1SectieStatistiek.getJSONObject("performance");
            JSONObject r1Overtredingen = r1SectieStatistiek.getJSONObject("overtredingen");
            JSONObject r1Snelheden = r1SectieStatistiek.getJSONObject("snelheden");
            JSONObject r1Passages = r1SectieStatistiek.getJSONObject("passages");

            /* Setting the PerformanceType */
            R1PerformanceType.setOvertredingenratio(Double.parseDouble(r1Performance.get("overtredingenratio").toString()));
            R1PerformanceType.setHandhaafratio(Double.parseDouble(r1Performance.get("handhaafratio").toString()));
            R1PerformanceType.setAutoratio(Double.parseDouble(r1Performance.get("autoratio").toString()));
            R1PerformanceType.setTijdvolledigbeschikbaar(r1Performance.get("tijdvolledigbeschikbaar").toString());
            R1PerformanceType.setMatchratio(Double.parseDouble(r1Performance.get("matchratio").toString()));
            
            /* Setting Up OvertredingenType */
            R1OvertredingenType.setAuto(Integer.parseInt(r1Overtredingen.get("auto").toString()));
            R1OvertredingenType.setOverigPardon(Integer.parseInt(r1Overtredingen.get("Overig-pardon").toString()));
            R1OvertredingenType.setOvertredingenTotaal(Integer.parseInt(r1Overtredingen.get("overtredingen-totaal").toString()));
            R1OvertredingenType.setDubbeleOvertredingenPardon(Integer.parseInt(r1Overtredingen.get("Dubbele-overtredingen-pardon").toString()));
            R1OvertredingenType.setHand(Integer.parseInt(r1Overtredingen.get("hand").toString()));
            
            /* Setting Up SnelhedenType */
            R1SnelhedenType.setGemiddeld(Double.parseDouble(r1Snelheden.getJSONObject("gemiddeld").get("content").toString()));
            R1SnelhedenType.setMax(Double.parseDouble(r1Snelheden.getJSONObject("max").get("content").toString()));
            
            /* Setting Up Passages */
            R1PassagesType.setTotalIn(Integer.parseInt(r1Passages.get("totaal-in").toString()));
            R1PassagesType.setTotalUit(Integer.parseInt(r1Passages.get("totaal-uit").toString()));
            
            /* Setting R1 */
            r1.setMatches(Integer.parseInt(r1SectieStatistiek.get("matches").toString()));
            r1.setOvertredingenType(R1OvertredingenType);
            r1.setPerformanceType(R1PerformanceType);
            r1.setSnelheden(R1SnelhedenType);
            r1.setpassagesType(R1PassagesType);
           
            
           
            /* Objects for holding L2 */
            L2 l2 = new L2();

            PerformanceType L2PerformanceType = new PerformanceType();
            OvertredingenType L2OvertredingenType = new OvertredingenType();
            SnelhedenType L2SnelhedenType = new SnelhedenType();
            PassagesType L2PassagesType = new PassagesType();

            /* Reading L - 2 Details */
            JSONObject l2SectieStatistiek = sectie.getJSONObject(1).getJSONObject("Sectie-statistiek");

            /* Extracting JSON subsections */
            JSONObject l2Performance = l2SectieStatistiek.getJSONObject("performance");
            JSONObject l2Overtredingen = l2SectieStatistiek.getJSONObject("overtredingen");
            JSONObject l2Snelheden = l2SectieStatistiek.getJSONObject("snelheden");
            JSONObject l2Passages = l2SectieStatistiek.getJSONObject("passages");

            /* Setting the PerformanceType */
            L2PerformanceType.setOvertredingenratio(Double.parseDouble(l2Performance.get("overtredingenratio").toString()));
            L2PerformanceType.setHandhaafratio(Double.parseDouble(l2Performance.get("handhaafratio").toString()));
            L2PerformanceType.setAutoratio(Double.parseDouble(l2Performance.get("autoratio").toString()));
            L2PerformanceType.setTijdvolledigbeschikbaar(l2Performance.get("tijdvolledigbeschikbaar").toString());
            L2PerformanceType.setMatchratio(Double.parseDouble(l2Performance.get("matchratio").toString()));
            
            /* Setting Up OvertredingenType */
            L2OvertredingenType.setAuto(Integer.parseInt(l2Overtredingen.get("auto").toString()));
            L2OvertredingenType.setOverigPardon(Integer.parseInt(l2Overtredingen.get("Overig-pardon").toString()));
            L2OvertredingenType.setOvertredingenTotaal(Integer.parseInt(l2Overtredingen.get("overtredingen-totaal").toString()));
            L2OvertredingenType.setDubbeleOvertredingenPardon(Integer.parseInt(l2Overtredingen.get("Dubbele-overtredingen-pardon").toString()));
            L2OvertredingenType.setHand(Integer.parseInt(l2Overtredingen.get("hand").toString()));
            
            /* Setting Up SnelhedenType */
            L2SnelhedenType.setGemiddeld(Double.parseDouble(l2Snelheden.getJSONObject("gemiddeld").get("content").toString()));
            L2SnelhedenType.setMax(Integer.parseInt(l2Snelheden.getJSONObject("max").get("content").toString()));
            
            /* Setting Up Passages */
            L2PassagesType.setTotalIn(Integer.parseInt(l2Passages.get("totaal-in").toString()));
            L2PassagesType.setTotalUit(Integer.parseInt(l2Passages.get("totaal-uit").toString()));
            
            /* Setting R1 */
            l2.setMatches(Integer.parseInt(l2SectieStatistiek.get("matches").toString()));
            l2.setOvertredingenType(L2OvertredingenType);
            l2.setPerformanceType(L2PerformanceType);
            l2.setSnelheden(L2SnelhedenType);
            l2.setpassagesType(L2PassagesType);
            
            /* Setting Sectie */
            mainSectie=new Sectie();
            mainSectie.setDate(date);
            mainSectie.setR1(r1);
            mainSectie.setL2(l2);
            
            
            
            //System.out.println(mainSectie);
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }

    }
    
    
    public Sectie getSectie() {
		return mainSectie;
	}

	public void setSectie(Sectie sectie) {
		this.mainSectie = sectie;
	}
}
