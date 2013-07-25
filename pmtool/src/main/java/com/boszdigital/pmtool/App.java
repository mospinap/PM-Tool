package com.boszdigital.pmtool;

import com.boszdigital.pmtool.service.TimeComparisonService;

/**
 * PM Tool [Project Manager Tool]
 * 
 * This tool is designed to support the processes of the Project Managers at
 * Bosz Digital.
 * 
 */
public class App {
	public static void main(String[] args) {
		String[] cad = new String[3];
		cad[0] = "/Users/daniel.hoyos/Desktop/";
		cad[1] = "staffoutput_0815-0821.csv";
		cad[2] = "fnd_gfm_0815-0821.tsv";
		TimeComparisonService timeComparisonService = new TimeComparisonService();
		timeComparisonService.run(cad);
	}
}
