# JiraConnectorForExcel
A basic 2-way JIRA-Excel communicator implemented using Excel VBA and JIRA REST APIs

## Supported Operations
| Field               |	Create  |	Query |	Update  |
|---------------------|---------|-------|---------|
| Summary	            | Yes     | Yes   |	Yes     |
| Assignee            | Yes     | Yes   |	Yes     |
| Status	            | No      | Yes   |	Yes     |
| Custom Field 0      | Yes     | No    |	No      |
| Fix Version         | Yes     | Yes   |	Yes     |
| Epic Key            | Yes     | No    |	No      |
| Priority            | Yes     | Yes   |	Yes     |
| Description         | Yes     | Yes   |	Yes     |
| Label	              | Yes     | No    |	No      |
| Custom Field 1      | Yes     | Yes   |	Yes     |
| Custom Field 2      | No      | Yes   |	Yes     |
| Component           | No      | Yes   |	Yes     |
| Due Date            | No      | Yes   |	Yes     |
| Start Date          | No      | Yes   |	Yes     |
| End Date            | No      | Yes   |	Yes     |
| Original Estimate   | No      | Yes   |	Yes     |
| Remaining Estimate  | No      | Yes   |	Yes     |
| Worklog	            | No      | Yes   |	Yes     |
| Comment	            | No      | Yes   |	Yes     |
| Epic Link	          | No      | Yes   |	Yes     |
| Blocked By	        | No      | Yes   |	Yes     |
| Blocks	            | No      | Yes   |	Yes     |

## Sheets
| Sheet Create              | Used to create issues in JIRA |
|---------------------------|-------------------------------|
| Get Template              | Populates  a list of 22 issues that are defined as a template. This can be used for new requirement tracking, technical readiness and customer requests.|
| Create Issues from Table	|Create JIRA items for the enteries specified in "*JiraCreateIssueTable*". **Tested only for issue type: Story**|

| Sheet Query_Update        | Used to query and update JIRA items based on certain criteria |
|---------------------------|-------------------------------|
|Supported queries          | IDSID search: Gets the IDSID of user name provided in cell D2 <br> Get issues with user provided text in cell B1<br>Get issues in filter ID provided in cell B2<br>Get all issues assigned to user specified in cell B3<br>Get issue ID mentioned in cell B4<br>Get issues linked to epic ID mentioned in cell B5<br>Cell B6 calls out the time for a requested operation (search/update)<br>Cell B7 tells the time of last JIRA API call. This is required to assess session key generation<br>Cell D1 holds the JIRA cookie key for current session<br>Cell D2 provides an option to query for IDSID from employee's full name|
|JiraQueryUpdateTable <br>(Table with entries populated from query as well as used to track and determine changes to be updated to JIRA)"|Changes are tracked with <span style="color:yellow">**Yellow**</span> cell color<br>This is accomplised using excel's sheet change library because of which *Undo functionality will not work in the table*<br>**Workaround: implement Undo stack for the table. This is currently not supported**<br>Another limitation is **bulk copy/paste traction is not handled**. It is recommended to use cell by cell updates instead of bulk or manually change the cell color to yellow<br>Deleting cell content changes cell color to **Red**. This is only for reference and is not updated to JIRA<br>On successful update to JIRA, cell color changes to **Green**<br>Failure to update to JIRA retains the cell's **Yellow** color. Error message will be captured in column S<br>**Red backfill**: Open items whose end date falls within 2 weeks of milestone release date. This means we have little room for error|
|Get Description	          | Macro to query as well as update description of JIRA issue selected in JiraQueryUpdateTable. Implemented using excel's userform|
|Get Comment	              | Macro to query, modify, delete and add new comments. Implemented using excel's userform|
|Get Worklog                | Macro to query, modify, delete and add new worklog. Implemented using excel's userform|
|Resume Tracking            |	Use this to resume modifications in *JiraQueryUpdateTable*. Tracking might get disabled in case of program crash|
|Update to JIRA	            | Reads all updates in *JiraQueryUpdateTable* and uses rest APIs to update JIRA|

| Sheet IDs        | Multiple tables that capture few of the field's supported values |
|------------------|-------------------------------|
|Refresh IDs       |	Clears all tables and queries the data from JIRA|
|JiraIssueTypeIDs  |	All supported JIRA issue types|
|JiraProjectIDs    |	All projects in JIRA|
|JiraStatusIDs     |	All states possible in JIRA (not tied to an issue type)|
|Proj1FixVersionIds|	Fix version list of project 1|
|Proj2FixVersionIds|	Fix version list of project 2|
|JiraFieldIds      |	All fiels in JIRA|
|FavouriteFilters  |	Favourite filter list of current user|
|ProjFlags         |	Flags of specified project|
|JiraPriorityIds   |	All priorities of JIRA|
|ProjComponents    |	All components of specified project|
|ProjPlatforms     |	All platforms of specified project|
|ProjOSes          |	All OSes of specified project|
|ProjUsers         |	All users assignable to an issue in specified project|
