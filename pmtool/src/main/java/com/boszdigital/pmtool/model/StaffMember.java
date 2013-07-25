package com.boszdigital.pmtool.model;

import java.util.HashMap;

/**
 * The representation of a member of the staff.
 * 
 * @author <a href="mailto:daniel.hoyos@boszdigital.com">Daniel Hoyos</a>
 *
 */

public class StaffMember {
	
	/** The code of the staff member generated from the name */
	private String codeName;
	
	/** The full name of the staff member */
	private String fullName;
	
	/** The projects assigned to the staff member */
	private HashMap<String, Project> projects;
	
	/** The total amount of time spent on projects */
	private float totalTime;
	
	/* -- Constructors -- */
	public StaffMember(){
		
	}
	
	public StaffMember(String codeName, String fullName){
		this.codeName = codeName;
		this.fullName = fullName;
		this.projects = new HashMap<String, Project>();
	}
	
	/* -- Getters and Setters -- */
	public String getCodeName() {
		return codeName;
	}

	public void setCodeName(String codeName) {
		this.codeName = codeName;
	}

	public String getFullName() {
		return fullName;
	}

	public void setFullName(String fullName) {
		this.fullName = fullName;
	}

	public HashMap<String, Project> getProjects() {
		return projects;
	}

	public void setProjects(HashMap<String, Project> projects) {
		this.projects = projects;
	}
	
	/* -- Other methods -- */
}
