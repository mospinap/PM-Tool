PM-Tool
=======

Project Manager Tool

The PM-Tool is an app designed to support Project Managers at Bosz Digital with their daily activities.

========

Version 0.0.1

This inital version consists on a single jar file with a single function: Compare the times tracket at ActiveTime and Oracle.
The system reads the two files and generates a report with the time differences. The files must be located on a folder defined by the user and given to the system by arguments.

========

Running

1. Download the app.
2. Unzip de downloaded file to a location of your choosing.
3. From the location the app was unziped, executed the following command:

	java -jar pmtool.jar {path/to/folder/} {active_time_file_name.ext} {oracle_file_name.ext}

			{path/to/folder/}			Is the path to the folder where the files are located and where
										the output is to be stored.
			{active_time_file_name.ext}	The name and extention of the active time file.
			{oracle_file_name.ext}		The name and extention of the oracle file.

