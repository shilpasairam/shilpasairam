### What is this repository for? ###

This repository contains the automation scripts related to LiveSLR application.

Framework: Python Pytest framework with Selenium

### How do I set up? ###

The setup for Python pytest automation framework

1. Prerequisites

	a) Install Python 3.x
	b) Add Python 3.x to your PATH environment variable
	c) If you do not have it already, get pip (NOTE: Most recent Python distributions come with pip)
	d) pip install -r requirements.txt to install dependencies

2. Refer cliText.txt file in the repository to get the commands to run the test case

3. Repository Structure

	Acutal Outputs: Saves the downloaded reports
	Configuration: For all configurations and credential files
	Logs: Log files for all tests
	Pages: Contains our Base Page, different Page Objects
	Reports: Contains screenshots and save the report.html file
	testCases: Actual testcases are present here
	Testdata: Contains Testdata required for the execution
	utilities: All utility modules (CustomLogger, LogScreenshot, ReadProperties) are kept in this folder
