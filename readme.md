
# Dimensions Analytics Pay Code Mapping Tool

This excel VBA code assists in mapping pay codes to analytics categories within UKG Dimensions. It works alongside UKG's excel pay code mapping workbook. The tool loops through a table that includes comma separated values of pay codes for each analytics category. From this loop it prepares a JSON string that will be accepted in an API call to the Dimensions tenant in order to either POST or PUT an analytics category into the Dimensions tenant. Outlined below are the steps taken.


1. Fill out UKG's Excel Pay Code Mapping Kit
2. Review the updated DDU Load Sheet after populating. It will look something similar to the below...

| ID  | Category Name| Category Description | Pay Code - Include| Pay Code - Exclude | API Response|
|-----|--------------|----------------------|-------------------|--------------------|-------------|
| 7   | Overtime     | Overtime Desc        | Daily Overtime, Overtime|              |             |
|     | Productive   | Productive Desc      | Callback, Charge Pay, Daily Overtime, Overtime, Regular | Charge Pay| |

3. Run the code for the DDU Load Table. When run, the tool will make the following decisions to build the JSON body and make the API call.
    1. If an ID exists, API call will be a PUT to update the mapping category as one with the same name and description already exists. Else, the tool will create an API POST call to create the category in the system.
    2. For each comma separated value of pay codes in the row, the tool will build a JSON string to attach to the larger JSON requst body that deciphers whether the paycode should be included in the category as COST ONLY or not.
    3. After each request is sent, the API response column will be populated with the respone message.
