package com.boszdigital.pmtool.model;

/**
 * The representation of a project.
 * 
 * @author <a href="mailto:daniel.hoyos@boszdigital.com">Daniel Hoyos</a>
 * 
 */
public class Project {

	/** The code assigned to the project */
	private String code;
	/** The name assigned to the project */
	private String name;
	/** The amount of time spent on the project */
	private float time;

	/* -- Constructors -- */
	public Project() {

	}

	public Project(String code, String name, float time) {
		this.code = code;
		this.name = name;
		this.time = time;
	}

	/* -- Getters and Setters -- */

	public String getCode() {
		return code;
	}

	public void setCode(String code) {
		this.code = code;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public float getTime() {
		return time;
	}

	public void setTime(float time) {
		this.time = time;
	}

	/* -- Other methods -- */

	/**
	 * This method adds hours to the total time of the project
	 * 
	 * @param hours
	 *            The amount of hours to add to the project
	 */
	public void addTime(float hours) {
		this.time = this.time + hours;
	}
}
