# Project
Multiple powershell scripts for IBM Qradar

# Table of content
1. [magic-qradar script](https://github.com/tacosaure/magic-qradar#magic-qradar-script)
2. [qradar_report_extraction_xml2csv script](https://github.com/tacosaure/magic-qradar#qradar_report_extraction_xml2csv-script)


## magic-qradar script
Powershell script to query Qradar and generate KPI

### Description
#### Overview
Powershell script allowing to query IBM Qradar easily in order to retrieve information or to generate KPI (Key Performance Indicators). 

The script is composed of multiple functions grouped by their use:
1. date/time handling functions
2. CSV and XLSX functions
3. IBM Qradar API functions
4. IBM Qradar data formating functions
5. KPI functions
6. KPI generation (output) function

#### Parameters
- qradar_api_key => token API generated within IBM Qradar
- start_date => "dd/MM/YYYY" format.  Start date of the search. Default = epochtime (01/01/1970)
- end_date => "dd/MM/YYYY" format. End date of the search. Default = now

#### Result
The function KPI_generation is called by the main script. It generates a XLSX file, named _kpi_qradar.xslx_, wich is composed of 11 worksheets:
- rule_and_BB_modified => list of rules and building block that have been created or modified during a period of time
- KPI_rule_and_BB_modified_status => numbers of rules and building block that have been created or modified during a period of time + bar chart
- closingReason_by_rule => numbers of closing reason of closed offenses by rules during a period of time
- closingReason_by_rule_stacked => numbers of closing reason of closed offenses by rules during a period of time + bar stacked chart
- KPI_offense_by_rule => number of offense triggered by rule during a period of time + pie chart
- offense_status
  - new => numbers of new offense triggered during a period of time
  - close => numbers of offenses that have been closed during a period of time
  - open => numbers of offense triggered during a period of time that are still open
  - Backlog open (total) => current backlog, offense still open (open + offense triggered before the specific period of time)
  - Total offense in Qradar
  - backlog_assigned => current backlog assigned (handled), offense still open but assigned (handled)
  - total_handled => backlog_assigned + close during a periode of time, number of offense handled during a period of time
- KPI_offense_backlog_by_rule => backlog (current open offenses) by rule + bar chart
- KPI_closingReason => Number of offenses by closing reasong during a period of time
- partial_offense_dataset => extract of offenses triggered during a period of time
- offense_full_dataset => extract of all offenses available in IBM Qradar
- rule_BB_dataset => extract of all rule and building blocks in IBM qradar

If you cannot install the importexcel powershell module, you can generate csv files for each sheets.

### Requirements
- Powershell
- Access to IBM Qradar API (API token)
- [importexcel powershell module](https://www.powershellgallery.com/packages/ImportExcel/7.1.0)

### First steps
We need to add the IBM Qradar API URL in the script.
```
$qradar_api_url = 'https://example_url_qradar.com/api/'
```

### How to use the script
To run the script, make sure you can reach your IBM Qradar console and execute the following command in powershell:
```
.\magic-qradar.ps1 -qradar_api_key 'XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXX' [-start_date "dd/MM/YYYY"] [-end_date "dd/MM/YYYY"]
```
## Warning
This script has been tested on IBM Qradar on CLOUD with the API version 14.0. Furthermore, the date format used was "dd/MM/YYYY", I do not know if there is an impact with computer using "MM/dd/YYYY" date format by default.

You may have warning messages by importexcel dealing with the closing Reason names that contains unsupported characters and have been converted into "\_" (closingReason_by_rule_stacked sheet). To solve this issue, it is a little bit tricky. No names refering cells (in a excel formula) can contains characters different than letters, numbers, "." and "\_". This is not a blocking point and you can skip these warnings.

### Need to be tested
- [ ] Get-SavedSearchDependentsTaskResults
- [ ] Get-SavedSearchDependentsTask
- [ ] Get-SavedSearchDependents

### Next improvements
- [ ] Add execution information
- [ ] Add verbose mode information
- [ ] Add a loading bar
- [ ] Add comments
- [ ] Create a report of objects dependency (BB, rule, saved searches, reference data ...)

### Steps while waiting for improvement implementation
#### BB dependencies (BB used by another rule/BB)
1. GET - /analytics/building_blocks => get the list of the BB ID
2. GET - /analytics/building_blocks/{id}/dependents => check dependencies and return the task ID for the following actions = id
3. GET - /analytics/rules/rule_dependent_tasks/{task_id} => if still running, we can have the number of the dependancies (rule, saved search, historical correllation...) = number_of_dependents,task_components,task_components[0].number_of_dependents (custom rules),task_components[1].number_of_dependents (search criteria),task_components[2].number_of_dependents (offense saved searches),task_components[3].number_of_dependents (historical correlation profiles),id,status
4. GET - /analytics/rules/rule_dependent_tasks/{task_id}/results => to get the name of the dependancies (name of the rule, saved search...) = dependent_id, dependent_name,dependent_type,dependent_owner,user_has_edit_permissions,blocking,dependent_group_ids

#### Rule dependencies (rule used by another rule/BB)
1. GET - /analytics/rules => get the list of the rule ID
2. GET - /analytics/rules/{id}/dependents => check dependencies and return the task ID for the following actions = id
3. GET - /analytics/rules/rule_dependent_tasks/{task_id} => if still running, we can have the number of the dependancies (rule, saved search, historical correllation...) = number_of_dependents,task_components,task_components[0].number_of_dependents (custom rules),task_components[1].number_of_dependents (search criteria),task_components[2].number_of_dependents (offense saved searches),task_components[3].number_of_dependents (historical correlation profiles),id,status
4. GET - /analytics/building_blocks/building_block_dependent_tasks/{task_id}/results => to get the name of the dependancies (name of the rule, saved search...) = dependent_id, dependent_name,dependent_type,dependent_owner,user_has_edit_permissions,blocking,dependent_group_ids

### Keywords
SIEM IBM Qradar KPI Automation SOC cyberanalyst reporting security operations center cybersecurity key performance indicators


2023-05-12 (ISO 8601)

## qradar_report_extraction_xml2csv script
Powershell script to retreive and convert basic/useful information from [qradar reports extraction files (xml)](https://www.ibm.com/docs/en/qsip/7.4?topic=content-exporting-all-custom-specific-type)

Usecase : Review all qradar reports in order to delete those not used anymore. These data may be used with magic-qradar script in order to get the dependency related to those reports

