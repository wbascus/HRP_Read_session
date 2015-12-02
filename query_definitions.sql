--Subunit_level_tracker
SELECT u3.*, 
s3.CountOffile_name AS S3_file_count, 
s3.LastOffile_name AS S3_most_recent_file, 
s3.FirstOfcontact AS S3_contact_info
FROM (SELECT u2.*, 
s2.CountOffile_name AS S2_file_count, 
s2.LastOffile_name AS S2_most_recent_file, 
s2.FirstOfcontact AS S2_contact_info 
FROM (SELECT u.unitID,
u.unit_name, 
u.sub_unit, 
u.unit_join, 
s1.countoffile_name AS S1_file_count, 
s1.LastOffile_name AS S1_most_recent_file, 
s1.FirstOfcontact AS s1_contact_info 
FROM units_summary AS u 
LEFT JOIN subunit_submittals_1 AS s1 ON u.unitID = s1.unitID)  AS u2 
LEFT JOIN subunit_submittals_2 AS s2 ON u2.unitID = s2.unitID)  AS u3 
LEFT JOIN subunit_submittals_3 AS s3 ON u3.unitID = s3.unitID;

--subunit_submittals X
SELECT s.session_no, 
s.unitID, 
Count(s.contact) AS Contact_ct, 
Count(s.truncated_file_name) AS CountOffile_name, 
Last(s.truncated_file_name) AS LastOffile_name, 
First(s.contact) AS FirstOfcontact
FROM submittal_summary s
GROUP BY s.session_no, 
s.unitID
HAVING (((s.session_no)="1"));

--Unit_level_tracker
SELECT m.unit_name, 
m.S1_file_count, 
m.S1_responses,
m.S2_file_count,
m.S2_responses, 
m.S3_file_count,
m.S3_responses
FROM (SELECT u3.*, 
s3.file_name_ct AS S3_file_count, 
s3.responses_sum AS S3_responses 
FROM (SELECT u2.*, 
s2.file_name_ct AS S2_file_count, 
s2.responses_sum AS S2_responses 
FROM (SELECT u.unit_name, 
s1.file_name_ct AS S1_file_count, 
s1.responses_sum AS S1_responses 
FROM units_summary AS u 
LEFT JOIN unit_submittals_1 AS s1 ON u.unit_name = s1.unit_name)  AS u2 
LEFT JOIN unit_submittals_2 AS s2 ON u2.unit_name = s2.unit_name)  AS u3 
LEFT JOIN unit_submittals_3 AS s3 ON u3.unit_name = s3.unit_name)  AS m
GROUP BY m.unit_name, 
m.S1_file_count, 
m.S2_file_count, 
m.S3_file_count, 
m.S1_responses,
m.S2_responses,
m.S3_responses;

--unit_sunmittals_X
SELECT s.session_no, 
s.unit_name, 
Count(s.truncated_file_name) AS file_name_ct, 
Sum(s.responses) AS responses_sum
FROM submittal_summary s
GROUP BY s.session_no, 
s.unit_name
HAVING (((s.session_no)="2"));

--Grand Union
SELECT DISTINCT * FROM (
SELECT u3.*, 
s3.[9A_2A], 
s3.[9A_3A], 
s3.[9B_2B], 
s3.[10_2A], 
s3.[10_3A], 
s3.[11A_2A], 
s3.[11A_3A], 
s3.[11A_3B], 
s3.[11A_4A], 
s3.[11B_2A], 
s3.[11B_3B], 
s3.[11B_3A], 
s3.[12_2A], 
s3.[12_3A], 
s3.[12_4A]
FROM (SELECT u2.*, 
s2.[5A_2B], 
s2.[5A_4A], 
s2.[5B_2B], 
s2.[5B_3A], 
s2.[5B_4A], 
s2.[6A_2A], 
s2.[6A_3A], 
s2.[6B_1A], 
s2.[6B_2A], 
s2.[6B_3A], 
s2.[7_4A], 
s2.[8A_1A], 
s2.[8A_2A], 
s2.[8A_2B], 
s2.[8A_3A], 
s2.[8B_1B], 
s2.[8B_2A], 
s2.[8B_3A], 
s2.[8B_4A], 
s2.[8C_2A], 
s2.[8C_3A] 
FROM (SELECT u.unitID, 
u.budget_no_corrected, 
u.EID_corrected,  
u.org_team_corrected_255,
s1.[1_2C], 
s1.[1_3B], 
s1.[2_2A],
s1.[4A_2A], 
s1.[4B_2A], 
s1.[4B_3A], 
s1.[4C_3A] 
FROM unique_employee_roles AS u 
LEFT JOIN session_1 AS s1 ON (u.org_team_corrected_255  = s1.org_team_corrected_255) AND (u.budget_no_corrected  = s1.budget_no_corrected) AND (u.EID_corrected  = s1.EID_corrected) AND (u.unitID  = s1.unitID))  AS u2 
LEFT JOIN session_2 AS s2 ON (u2.org_team_corrected_255 = s2.org_team_corrected_255) AND (u2.budget_no_corrected = s2.budget_no_corrected) AND (u2.EID_corrected = s2.EID_corrected) AND (u2.unitID = s2.unitID))  AS u3 
LEFT JOIN session_3 AS s3 ON (u3.org_team_corrected_255 = s3.org_team_corrected_255) AND (u3.budget_no_corrected = s3.budget_no_corrected) AND (u3.EID_corrected = s3.EID_corrected) AND (u3.unitID = s3.unitID)
) m;

--Session 1
SELECT s.unitID, 
responses.EID_corrected, 
responses.org_team_mapID, 
responses.budget_no_corrected, 
responses.org_team_corrrected,
responses.[1_2C], 
responses.[1_3B], 
responses.[2_2A], 
responses.[4A_2A], 
responses.[4B_2A], 
responses.[4B_3A], 
responses.[4C_3A]
FROM submittal_summary s INNER JOIN responses ON s.submittalID = responses.submittalID
WHERE (((responses.EID_corrected)<>"") AND ((s.session_no)="1"));

--Session 2
SELECT s.unitID, 
responses.EID_corrected, 
responses.org_team_mapID, 
responses.budget_no_corrected,
responses.org_team_corrected,
responses.[5A_2B], 
responses.[5A_4A], 
responses.[5B_2B], 
responses.[5B_3A], 
responses.[5B_4A], 
responses.[6A_2A], 
responses.[6A_3A], 
responses.[6B_1A], 
responses.[6B_2A], 
responses.[6B_3A], 
responses.[7_4A], 
responses.[8A_1A], 
responses.[8A_2A], 
responses.[8A_2B], 
responses.[8A_3A], 
responses.[8B_1B], 
responses.[8B_2A], 
responses.[8B_3A], 
responses.[8B_4A], 
responses.[8C_2A], 
responses.[8C_3A]
FROM submittal_summary  s INNER JOIN responses ON s.submittalID = responses.submittalID
WHERE (((responses.EID_corrected)<>"") AND ((s.session_no)="2"));

--Session 3
SELECT s.unitID, 
responses.EID_corrected, 
responses.org_team_mapID, 
responses.budget_no_corrected,
responses.org_team_corrected,
responses.[9A_2A], 
responses.[9A_3A], 
responses.[9B_2B], 
responses.[10_2A], 
responses.[10_3A], 
responses.[11A_2A], 
responses.[11A_3A], 
responses.[11A_3B], 
responses.[11A_4A], 
responses.[11B_2A], 
responses.[11B_3A], 
responses.[11B_3B], 
responses.[12_2A], 
responses.[12_3A], 
responses.[12_4A], 
responses.date_recorded
FROM submittal_summary s INNER JOIN responses ON s.submittalID = responses.submittalID
WHERE (((responses.EID_corrected)<>"") AND ((s.session_no)="3"));

--New Role Mapping by field in order of role
SELECT g.unitID,
g.unit_join AS Unit, 
g.unit_cm AS [Change Manager],
g.budget_no_corrected AS [Home Dpt Budget Number], 
g.EID_corrected AS [EID], 
g.dw_name_first AS [Employee First Name],
g.dw_name_last AS [Employee Last Name],
g.org_team_corrected AS [Supervisory Org], 
IIf(g.[7_4A]> 0, "x","") AS [7_4A], 
IIf(g.[4A_2A]> 0, "x","") AS [4A_2A], 
IIf(g.[4B_2A] > 0, "x","") AS [4B_2A], 
IIf(g.[4B_3A] > 0, "x","") AS [4B_3A], 
IIf(g.[4C_3A] > 0, "x","") AS [4C_3A], 
IIf(g.[8A_2B] > 0, "x","") AS [8A_2B], 
IIf(g.[5B_3A] > 0, "x","") AS [5B_3A], 
IIf(g.[6B_2A] > 0, "x","") AS [6B_2A], 
IIf(g.[8B_2A] > 0, "x","") AS [8B_2A], 
IIf(g.[11A_3A] > 0, "x","") AS [11A_3A], 
IIf(g.[11A_3B] > 0, "x","") AS [11A_3B], 
IIf(g.[11B_3A] > 0, "x","") AS [11B_3A], 
IIf(g.[11B_3B] > 0, "x","") AS [11B_3B], 
IIf(g.[10_3A] > 0, "x","") AS [10_3A], 
IIf(g.[6A_3A] > 0, "x","") AS [6A_3A], 
IIf(g.[6B_3A] > 0, "x","") AS [6B_3A], 
IIf(g.[12_4A] > 0, "x","") AS [12_4A], 
IIf(g.[5A_2B] > 0, "x","") AS [5A_2B], 
IIf(g.[5A_4A] > 0, "x","") AS [5A_4A], 
IIf(g.[5B_2B] > 0, "x","") AS [5B_2B], 
IIf(g.[5B_4A] > 0, "x","") AS [5B_4A], 
IIf(g.[6B_1A] > 0, "x","") AS [6B_1A], 
IIf(g.[8A_1A] > 0, "x","") AS [8A_1A], 
IIf(g.[8A_3A] > 0, "x","") AS [8A_3A], 
IIf(g.[8B_1B] > 0, "x","") AS [8B_1B], 
IIf(g.[8B_3A] > 0, "x","") AS [8B_3A], 
IIf(g.[8b_4A] > 0, "x","") AS [8b_4A], 
IIf(g.[8C_3A] > 0, "x","") AS [8C_3A], 
IIf(g.[9A_2A] > 0, "x","") AS [9A_2A], 
IIf(g.[9B_2B] > 0, "x","") AS [9B_2B], 
IIf(g.[10_2A] > 0, "x","") AS [10_2A], 
IIf(g.[11A_2A] > 0, "x","") AS [11A_2A], 
IIf(g.[11A_4A] > 0, "x","") AS [11A_4A], 
IIf(g.[11B_2A] > 0, "x","") AS [11B_2A], 
IIf(g.[12_2A] > 0, "x","") AS [12_2A], 
IIf(g.[6A_2A] > 0, "x","") AS [6A_2A], 
IIf(g.[8A_2A] > 0, "x","") AS [8A_2A], 
IIf(g.[8C_2A] > 0, "x","") AS [8C_2A], 
IIf(g.[9A_3A] > 0, "x","") AS [9A_3A], 
IIf(g.[12_3A] > 0, "x","") AS [12_3A], 
IIf(g.[1_2C] > 0, "x","") AS [1_2C], 
IIf(g.[1_3B] > 0, "x","") AS [1_3B], 
IIf(g.[2_2A] > 0, "x","") AS [2_2A]
FROM [New Grand Union] AS g
ORDER by g.unitID, g.org_team_corrected, g.budget_no_corrected, g.dw_name_last;

--New Role_Mapping_by_field_in_order_of_scenario
SELECT g.unitID,
g.unit_join AS Unit, 
g.unit_cm AS [Change Manager],
g.budget_no_corrected AS [Home Dpt Budget Number], 
g.EID_corrected AS [EID], 
g.dw_name_first AS [Employee First Name],
g.dw_name_last AS [Employee Last Name],
g.org_team_corrected AS [Supervisory Org], 
IIF(g.[1_2C] > 0, "x","") AS [1_2C], 
IIF(g.[1_3B] > 0, "x","") AS [1_3B], 
IIF(g.[2_2A] > 0, "x","") AS [2_2A],
IIF(g.[4A_2A] > 0, "x","") AS [4A_2A], 
IIF(g.[4B_2A] > 0, "x","") AS [4B_2A], 
IIF(g.[4B_3A] > 0, "x","") AS [4B_3A], 
IIF(g.[4C_3A] > 0, "x","") AS [4C_3A], 
IIF(g.[5A_2B] > 0, "x","") AS [5A_2B], 
IIF(g.[5A_4A] > 0, "x","") AS [5A_4A], 
IIF(g.[5B_2B] > 0, "x","") AS [5B_2B], 
IIF(g.[5B_3A] > 0, "x","") AS [5B_3A], 
IIF(g.[5B_4A] > 0, "x","") AS [5B_4A], 
IIF(g.[6A_2A] > 0, "x","") AS [6A_2A], 
IIF(g.[6A_3A] > 0, "x","") AS [6A_3A], 
IIF(g.[6B_1A] > 0, "x","") AS [6B_1A], 
IIF(g.[6B_2A] > 0, "x","") AS [6B_2A], 
IIF(g.[6B_3A] > 0, "x","") AS [6B_3A], 
IIF(g.[7_4A] > 0, "x","") AS [7_4A], 
IIF(g.[8A_1A] > 0, "x","") AS [8A_1A], 
IIF(g.[8A_2A] > 0, "x","") AS [8A_2A], 
IIF(g.[8A_2B] > 0, "x","") AS [8A_2B], 
IIF(g.[8A_3A] > 0, "x","") AS [8A_3A], 
IIF(g.[8B_1B] > 0, "x","") AS [8B_1B], 
IIF(g.[8B_2A] > 0, "x","") AS [8B_2A], 
IIF(g.[8B_3A] > 0, "x","") AS [8B_3A], 
IIF(g.[8B_4A] > 0, "x","") AS [8B_4A], 
IIF(g.[8C_2A] > 0, "x","") AS [8C_2A], 
IIF(g.[8C_3A] > 0, "x","") AS [8C_3A], 
IIF(g.[9A_2A] > 0, "x","") AS [9A_2A], 
IIF(g.[9A_3A] > 0, "x","") AS [9A_3A], 
IIF(g.[9B_2B] > 0, "x","") AS [9B_2B], 
IIF(g.[10_2A] > 0, "x","") AS [10_2A], 
IIF(g.[10_3A] > 0, "x","") AS [10_3A], 
IIF(g.[11A_2A] > 0, "x","") AS [11A_2A], 
IIF(g.[11A_3A] > 0, "x","") AS [11A_3A],
IIF(g.[11A_3B] > 0, "x","") AS [11A_3B],  
IIF(g.[11A_4A] > 0, "x","") AS [11A_4A], 
IIF(g.[11B_2A] > 0, "x","") AS [11B_2A], 
IIF(g.[11B_3A] > 0, "x","") AS [11B_3A], 
IIF(g.[11B_3B] > 0, "x","") AS [11B_3B],
IIF(g.[12_2A] > 0, "x","") AS [12_2A], 
IIF(g.[12_3A] > 0, "x","") AS [12_3A], 
IIF(g.[12_4A] > 0, "x","") AS [12_4A] 
FROM [New Grand Union] AS g 
ORDER by g.unitID, g.org_team_corrected, g.budget_no_corrected, g.dw_name_last;

--New Workday Role_Mapping_By Role_base
SELECT g.unitID, 
g.unit_join,
g.unit_cm,
g.budget_no_corrected, 
g.dw_name_first, 
g.dw_name_last, 
g.EID_corrected, 
g.org_team_corrected, 
IIf(g.[7_4A] > 0,"x","") AS I9, 
IIF((g.[4A_2A]+[g].[4B_2A]+[g].[4B_3A]+[g].[4C_3A]+[g].[8A_2B])> 0,"x","") AS ABP, 
IIf(([g].[5B_3A] + [g].[6B_2A] + [g].[8B_2A] + [g].[11A_3A] + [g].[11B_3A])>0,"x","") AS ACP, 
IIf(([g].[10_3A])>0,"x","") AS CP, 
IIf(([g].[6A_3A] + [g].[6B_3A] + [g].[12_4A])>0,"x","") AS CAC, 
IIf(([g].[5A_2B] + [g].[5A_4A] + [g].[5B_2B] + [g].[5B_4A] + [g].[6B_1A] + [g].[8A_1A] + [g].[8A_3A] + [g].[8B_1B] + [g].[8B_3A] + [g].[8b_4A] + [g].[8C_3A] + [g].[9A_2A] + [g].[9B_2B] + [g].[10_2A] + [g].[11A_2A] + [g].[11A_4A] + [g].[11B_2A] + [g].[12_2A])>0,"x","") AS HRC, 
IIf(([g].[6A_2A] + [g].[8A_2A] + [g].[8C_3A] + [g].[9A_3A] + [g].[12_3A])>0,"x","") AS HRP, 
IIf(([g].[1_2C] + [g].[1_3B] + [g].[2_2a])>0,"x","") AS TC 
FROM [New Grand Union] AS g

--New Workday_Role_Mapping_By_Role
SELECT b.unitID,
b.unit_join AS Unit, 
b.unit_cm AS [Change Manager],
b.budget_no_corrected AS [Home Dpt Budget Number], 
b.[dw_name_first] AS [Employee First Name], 
b.[dw_name_last] AS [Employee Last Name], 
b.EID_corrected AS EID, 
b.org_team_corrected AS [Supervisory Org], 
IIf(b.ABP="x","x","") AS ABP, 
"" AS AC, 
"" AS AD, 
"" AS AE, 
IIf([b].[ACP]="x","x","") AS ACP, 
IIf([b].[CP]="x","x","") AS CP, 
IIf([b].[CAC]="x","x","") AS CAC, 
IIf(([b].[HRP]="x" OR [b].[ACP] = "x"),"",IIf([b].[HRC]="x","x","")) AS HRC, 
"" as HRE,
IIf([b].[ACP]="x","x",IIf([b].[HRP]="x","x","")) AS HRP, 
IIf([b].[I9]="x","x","") AS I9, 
IIf([b].[ACP]="x","x",IIf([b].[HRP]="x","x","")) AS RP, 
IIf([b].[TC]="x","x","") AS TC
FROM [New_Workday_Role_Mapping_by_role_base] AS b;

--New Workday_Role_Mapping_By_role_transpose
SELECT [b].[Employee First Name] & " " & [b].[Employee Last Name] & " " & [b].[EID] AS [Employee Slug], 
b.Unit, 
b.ABP, 
"" AS AC, 
"" AS AD, 
"" AS AE, 
b.ACP, 
b.CP, 
b.CAC, 
b.HRC, 
"" AS HRE,
b.HRP, 
b.I9, 
b.RP, 
b.TC
FROM [New Workday_Role_Mapping_by_role] b;

--submittal_summary
SELECT t2.submittalID,
t2.date_submitted,
t2.contact,
t2.date_recorded,
t2.session_no,
t2.truncated_file_name,
t2.unitID,
COUNT(r.responseID) As EIDs
FROM (SELECT s.submittalID, 
s.date_submitted, 
s.contact,
s.date_recorded, 
s.session_no, 
s.truncated_file_name,
s.unitID
FROM submittals AS s
WHERE s.superceded = FALSE) t2
LEFT OUTER JOIN responses r 
ON t2.submittalID = r.submittalID
GROUP BY 
t2.submittalID,
t2.date_submitted,
t2.contact,
t2.date_recorded,
t2.session_no,
t2.truncated_file_name,
t2.unitID;

--New Grand Union
SELECT
t1.[gu.unitID] AS unitID,
t1.EID_corrected,
t1.dw_name_first,
t1.dw_name_last,
t1.budget_no_corrected,
t1.org_team_corrected,
us.unit_cm,
us.unit_join,
Count(t1.[1_2C]) AS [1_2C], 
Count(t1.[1_3B]) AS [1_3B],
Count(t1.[2_2A]) AS [2_2A],
Count(t1.[4A_2A]) AS [4A_2A],
Count(t1.[4B_2A]) AS [4B_2A],
Count(t1.[4B_3A]) AS [4B_3A],
Count(t1.[4C_3A]) AS [4C_3A],
Count(t1.[5A_2B]) AS [5A_2B],
Count(t1.[5A_4A]) AS [5A_4A],
Count(t1.[5B_2B]) AS [5B_2B],
Count(t1.[5B_3A]) AS [5B_3A],
Count(t1.[5B_4A]) AS [5B_4A],
Count(t1.[6A_2A]) AS [6A_2A],
Count(t1.[6A_3A]) AS [6A_3A],
Count(t1.[6B_1A]) AS [6B_1A],
Count(t1.[6B_2A]) AS [6B_2A],
Count(t1.[6B_3A]) AS [6B_3A],
Count(t1.[7_4A]) AS [7_4A],
Count(t1.[8A_1A]) AS [8A_1A],
Count(t1.[8A_2A]) AS [8A_2A],
Count(t1.[8A_2B]) AS [8A_2B],
Count(t1.[8A_3A]) AS [8A_3A],
Count(t1.[8B_1B]) AS [8B_1B],
Count(t1.[8B_2A]) AS [8B_2A],
Count(t1.[8B_3A]) AS [8B_3A],
Count(t1.[8B_4A]) AS [8B_4A],
Count(t1.[8C_2A]) AS [8C_2A],
Count(t1.[8C_3A]) AS [8C_3A],
Count(t1.[9A_2A]) AS [9A_2A],
Count(t1.[9A_3A]) AS [9A_3A],
Count(t1.[9B_2B]) AS [9B_2B],
Count(t1.[10_2A]) AS [10_2A],
Count(t1.[10_3A]) AS [10_3A],
Count(t1.[11A_2A]) AS [11A_2A],
Count(t1.[11A_3A]) AS [11A_3A],
Count(t1.[11A_3B]) AS [11A_3B],
Count(t1.[11A_4A]) AS [11A_4A],
Count(t1.[11B_2A]) AS [11B_2A],
Count(t1.[11B_3A]) AS [11B_3A],
Count(t1.[11B_3B]) AS [11B_3B],
Count(t1.[12_2A]) AS [12_2A],
Count(t1.[12_3A]) AS [12_3A],
Count(t1.[12_4A]) AS [12_4A]
FROM (SELECT * 
FROM (SELECT * FROM submittals s, responses r where r.submittalID = s.SubmittalID 
AND s.superceded = FALSE) gu
LEFT JOIN employees e 
ON e.dw_eid = gu.eid_corrected) t1,
units_summary us 
WHERE t1.[gu.unitID] = us.unitID
AND t1.EID_corrected <> ""
AND t1.EID_corrected <> "-"
GROUP BY t1.[gu.unitid], t1.org_team_corrected, t1.EID_corrected, t1.budget_no_corrected, t1.dw_name_first, t1.dw_name_last, us.unit_join, us.unit_cm
ORDER by t1.[gu.unitid], t1.dw_name_last, org_team_corrected, budget_no_corrected




