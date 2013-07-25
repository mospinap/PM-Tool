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
	public StaffMember() {

	}

	public StaffMember(String codeName, String fullName) {
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

	/**
	 * The method adds hours to an existing project assigned to the staff
	 * member. If the project does not exists, the method creates the project
	 * and adds the hours.
	 * 
	 * @param projectCode
	 *            The code of the project
	 * @param projectName
	 *            The name of the project
	 * @param hours
	 *            The amount of hours to add to the project
	 */
	public void addTime(String projectCode, String projectName, float hours) {
		Project project;
		if (projects.containsKey(projectCode)) {
			project = projects.get(projectCode);
			project.addTime(hours);

		} else {
			project = new Project(projectCode, projectName, hours);
		}
		projects.put(projectCode, project);
	}

	/**
	 * This method adds hours to the total time spent by the staff member.
	 * 
	 * @param hours
	 *            The amount of hours to add to the total time.
	 */
	public void addTotalTime(float hours) {
		this.totalTime = this.totalTime + hours;
	}
}
