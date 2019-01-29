Attribute VB_Name = "PopulateIssueCreationTemplate"
Sub PopulateIssueTemplate()
    Dim epic_key As String
    epic_key = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B2").Type2ue
    template(0, 1) = "Type1 Analysis & Effort Estimation"
    template(0, 2) = owner_1
    template(0, 3) = "Type1"
    template(0, 5) = epic_key
    template(0, 6) = "P1-Stopper"
    template(0, 7) = "||Task||Owner||Effort (days)||Can start date||Comments||\n|Design| | | | |\n|Coding| | | | |\n|UTF Development| | | | |\n|ULT| | | | |\n|Type1 review| | | | |\n|Bug fixes| | | | |"
    
    template(1, 1) = "Type2 Analysis & Effort Estimation"
    template(1, 2) = owner_2
    template(1, 3) = "Type2"
    template(1, 5) = epic_key
    template(1, 6) = "P1-Stopper"
    template(1, 7) = "||Task||Owner||Effort (days)||Can start date||Comments||\n|Design Spec| | | | |\n|Type3 Development| | | | |\n|Type3 review| | | | |\n|Type3 ULT| | | | |\n|Type3 fixes| | | | |"
    
    template(2, 1) = "Type5 Effort Estimation"
    template(2, 2) = owner_5
    template(2, 3) = "Type2"
    template(2, 5) = epic_key
    template(2, 6) = "P1-Stopper"
    template(2, 7) = "||Task||Owner||Effort (days)||Can start date||Comments||\n|Type5 round 1| | | | |\n|Type5 round 2| | | | |"
    
    template(3, 1) = "Schedule"
    template(3, 2) = owner_3
    template(3, 3) = "Type1,Type2"
    template(3, 5) = epic_key
    template(3, 6) = "P1-Stopper"
    template(3, 7) = "||Task||Owner||Effort (days)||Bandwidth (%)||Start date (WWxx.x)||End date (WWxx.x)||Comments||\n|Type4 Enabling| | | | | | |\n|Type1 Design| | | | | | |\n|Type1 Coding| | | | | | |\n|UTF Development| | | | | | |\n|Type1 ULT| | | | | | |\n|Type1 Type1 review| | | | | | |\n|Design Spec| | | | | | |\n|Type3 Development| | | | | | |\n|Type3 review| | | | | | |\n|Type3 ULT| | | | | | |\n|Type3 fixes| | | | | | |\n|Type3 qualification| | | | | | |\n|Type3 depolyment| | | | | | |\n|Type5 round 1| | | | | | |\n|Bug fixes| | | | | | |\n|Type5 round 2| | | | | | |\n|Mainline merge| | | | | | |"
    
    template(4, 1) = "Type1 Design"
    template(4, 2) = owner_1
    template(4, 3) = "Type1"
    template(4, 5) = epic_key
    template(4, 6) = "P2-High"
    template(4, 7) = ""
    
    template(5, 1) = "Type1 Design Review"
    template(5, 2) = owner_1
    template(5, 3) = "Type1"
    template(5, 5) = epic_key
    template(5, 6) = "P2-High"
    template(5, 7) = ""
    
    template(6, 1) = "Type1 Coding"
    template(6, 2) = owner_1
    template(6, 3) = "Type1"
    template(6, 5) = epic_key
    template(6, 6) = "P2-High"
    template(6, 7) = ""
    
    template(7, 1) = "Type1 UTF Development"
    template(7, 2) = owner_1
    template(7, 3) = "Type1"
    template(7, 5) = epic_key
    template(7, 6) = "P2-High"
    template(7, 7) = ""
    
    template(8, 1) = "Type1 ULT"
    template(8, 2) = owner_1
    template(8, 3) = "Type1"
    template(8, 5) = epic_key
    template(8, 6) = "P2-High"
    template(8, 7) = ""
    
    template(9, 1) = "Type1 Type1 Review"
    template(9, 2) = owner_1
    template(9, 3) = "Type1"
    template(9, 5) = epic_key
    template(9, 6) = "P2-High"
    template(9, 7) = ""
    
    template(10, 1) = "Type1 Bug Fixes"
    template(10, 2) = owner_1
    template(10, 3) = "Type1"
    template(10, 5) = epic_key
    template(10, 6) = "P2-High"
    template(10, 7) = ""
    
    template(11, 1) = "Type3 Design"
    template(11, 2) = owner_2
    template(11, 3) = "Type2"
    template(11, 5) = epic_key
    template(11, 6) = "P2-High"
    template(11, 7) = ""
    
    template(12, 1) = "Type3 Design Review"
    template(12, 2) = owner_2
    template(12, 3) = "Type2"
    template(12, 5) = epic_key
    template(12, 6) = "P2-High"
    template(12, 7) = ""
    
    template(13, 1) = "Type3 Development"
    template(13, 2) = owner_2
    template(13, 3) = "Type2"
    template(13, 5) = epic_key
    template(13, 6) = "P2-High"
    template(13, 7) = ""
    
    template(14, 1) = "Type3 ULT"
    template(14, 2) = owner_2
    template(14, 3) = "Type2"
    template(14, 5) = epic_key
    template(14, 6) = "P2-High"
    template(14, 7) = ""
    
    template(15, 1) = "Type3 Review"
    template(15, 2) = owner_2
    template(15, 3) = "Type2"
    template(15, 5) = epic_key
    template(15, 6) = "P2-High"
    template(15, 7) = ""
    
    template(16, 1) = "Type3 Fixes"
    template(16, 2) = owner_2
    template(16, 3) = "Type2"
    template(16, 5) = epic_key
    template(16, 6) = "P2-High"
    template(16, 7) = ""
    
    template(17, 1) = "Type3 Qualification"
    template(17, 2) = owner_2
    template(17, 3) = "Type2"
    template(17, 5) = epic_key
    template(17, 6) = "P2-High"
    template(17, 7) = ""
    
    template(18, 1) = "Type3 Deployment"
    template(18, 2) = owner_2
    template(18, 3) = "Type2"
    template(18, 5) = epic_key
    template(18, 6) = "P2-High"
    template(18, 7) = ""
    
    template(19, 1) = "Type5 Round 1"
    template(19, 2) = owner_5
    template(19, 3) = "Type2"
    template(19, 5) = epic_key
    template(19, 6) = "P2-High"
    template(19, 7) = ""
    
    template(20, 1) = "Type5 Round 2"
    template(20, 2) = owner_5
    template(20, 3) = "Type2"
    template(20, 5) = epic_key
    template(20, 6) = "P2-High"
    template(20, 7) = ""
    
    template(21, 1) = "Type4 Enabling"
    template(21, 2) = owner_4
    template(21, 3) = "Type1,Type2"
    template(21, 5) = epic_key
    template(21, 6) = "P2-High"
    template(21, 7) = ""
End Sub
