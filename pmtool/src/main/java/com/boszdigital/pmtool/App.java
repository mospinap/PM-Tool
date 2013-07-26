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
		TimeComparisonService timeComparisonService = new TimeComparisonService();
		timeComparisonService.run(args);
	}
}
