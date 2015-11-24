'Subunit_level_tracker'
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

'subunit_submittals X'
SELECT s.session_no, 
s.unitID, 
Count(s.contact) AS Contact_ct, 
Count(s.file_name) AS CountOffile_name, 
Last(s.file_name) AS LastOffile_name, 
First(s.contact) AS FirstOfcontact
FROM submittal_summary s
GROUP BY s.session_no, 
s.unitID
HAVING (((s.session_no)="1"));

'Unit_level_tracker'
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

'unit_sunmittals_X'
SELECT s.session_no, 
s.unit_name, 
Count(s.file_name) AS file_name_ct, 
Sum(s.responses) AS responses_sum
FROM submittal_summary s
GROUP BY s.session_no, 
s.unit_name
HAVING (((s.session_no)="2"));

'Grand Union'
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
u.unitID,
u.budget_no, 
u.first_name, 
u.last_name, 
u.EID, 
"" AS ["Supervisory Org"], 
u.org_team_mapID, 
s1.[1_2C], 
s1.[1_3B], 
s1.[2_2A],
s1.[4A_2A], 
s1.[4B_2A], 
s1.[4B_3A], 
s1.[4C_3A] 
FROM unique_employee_roles AS u 
LEFT JOIN session_1 AS s1 ON (u.org_team_mapID= s1.org_team_mapID) AND (u.budget_no = s1.budget_no) AND (u.EID = s1.EID) AND (u.unitID = s1.unitID))  AS u2 
LEFT JOIN session_2 AS s2 ON (u2.org_team_mapID= s2.org_team_mapID) AND (u2.budget_no = s2.budget_no) AND (u2.EID = s2.EID) AND (u2.unitID = s2.unitID))  AS u3 
LEFT JOIN session_3 AS s3 ON (u3.EID = s3.EID) AND (u3.budget_no = s3.budget_no) AND (u3.org_team_mapID = s3.org_team_mapID) AND (u3.unitID = s3.unitID);

'Session 1'
SELECT responses.responseID, 
s.submittalID, 
s.file_name, 
s.unitID, 
s.unit_join, 
responses.first_name, 
responses.last_name, 
responses.eid, 
responses.org_team_mapID, 
responses.budget_no, 
s.session_no, 
responses.[1_2C], 
responses.[1_3B], 
responses.[2_2A], 
responses.[4A_2A], 
responses.[4B_2A], 
responses.[4B_3A], 
responses.[4C_3A], 
responses.date_recorded,
s.unitID
FROM submittal_summary s INNER JOIN responses ON s.submittalID = responses.submittalID
WHERE (((responses.eid)<>"") AND ((s.session_no)="1"));

'Session 2'
SELECT responses.responseID, 
s.submittalID, 
s.file_name, 
s.unit_join, 
responses.first_name, 
responses.last_name, 
responses.EID, 
responses.org_team_mapID, 
responses.budget_no, 
s.session_no, 
responses.date_recorded, 
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
responses.[8C_3A], 
s.unitID
FROM submittal_summary  s INNER JOIN responses ON s.submittalID = responses.submittalID
WHERE (((responses.EID)<>"") AND ((s.session_no)="2"));

'Session 3'
SELECT responses.responseID, 
s.submittalID, 
s.file_name, 
s.unit_join, 
responses.first_name, 
responses.last_name, responses.EID, 
responses.org_team_mapID, 
responses.budget_no, 
s.session_no, 
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
responses.date_recorded, 
s.unitID
FROM submittal_summary s INNER JOIN responses ON s.submittalID = responses.submittalID
WHERE (((responses.EID)<>"") AND ((s.session_no)="3"));

'Role Mapping by field in order of role'
SELECT DISTINCT t2.unit_join AS Unit, 
t2.budget_no AS [Home Dpt Budget Number], 
t2.first_name AS [Employee First Name], 
t2.last_name AS [Employee Last Name], 
t2.EID, 
t2.org_team_corrected AS [Supervisory Org], 
t2.[7_4A], 
t2.[4A_2A], 
t2.[4B_2A], 
t2.[4B_3A], 
t2.[4C_3A], 
t2.[8A_2B], 
t2.[5B_3A], 
t2.[6B_2A], 
t2.[8B_2A], 
t2.[11A_3A], 
t2.[11A_3B], 
t2.[11B_3A], 
t2.[11B_3B], 
t2.[10_3A], 
t2.[6A_3A], 
t2.[6B_3A], 
t2.[12_4A], 
t2.[5A_2B], 
t2.[5A_4A], 
t2.[5B_2B], 
t2.[5B_4A], 
t2.[6B_1A], 
t2.[8A_1A], 
t2.[8A_3A], 
t2.[8B_1B], 
t2.[8B_3A], 
t2.[8b_4A], 
t2.[8C_3A], 
t2.[9A_2A], 
t2.[9B_2B], 
t2.[10_2A], 
t2.[11A_2A], 
t2.[11A_4A], 
t2.[11B_2A], 
t2.[12_2A], 
t2.[6A_2A], 
t2.[8A_2A], 
t2.[8C_2A], 
t2.[9A_3A], 
t2.[12_3A], 
t2.[1_2C], 
t2.[1_3B], 
t2.[2_2A]
FROM (SELECT * FROM (SELECT * FROM GrandUnion AS g INNER JOIN org_team_map AS s ON g.org_team_mapID = s.org_team_mapID)  AS t1 INNER JOIN units_summary AS u ON t1.unitID = u.unitID)  AS t2;

'Role_Mapping_by_field_in_order_of_scenario'
SELECT DISTINCT t2.unit_join AS Unit, 
t2.budget_no AS [Home Dpt Budget Number], 
t2.first_name AS [Employee First Name], 
t2.last_name AS [Employee Last Name], 
t2.EID, 
t2.org_team_corrected AS [Supervisory Org], 
t2.[1_2C], 
t2.[1_3B], 
t2.[2_2A],
 t2.[4A_2A], 
t2.[4B_2A], 
t2.[4B_3A], 
t2.[4C_3A], 
t2.[5A_2B], 
t2.[5A_4A], 
t2.[5B_2B], 
t2.[5B_3A], 
t2.[5B_4A], 
t2.[6A_2A], 
t2.[6A_3A], 
t2.[6B_1A], 
t2.[6B_2A], 
t2.[6B_3A], 
t2.[7_4A], 
t2.[8A_1A], 
t2.[8A_2A], 
t2.[8A_2B], 
t2.[8A_3A], 
t2.[8B_1B], 
t2.[8B_2A], 
t2.[8B_3A], 
t2.[8B_4A], 
t2.[8C_2A], 
t2.[8C_3A], 
t2.[9A_2A], 
t2.[9A_3A], 
t2.[9B_2B], 
t2.[10_2A], 
t2.[10_3A], 
t2.[11A_2A], 
t2.[11A_3A], 
t2.[11A_4A], 
t2.[11B_2A], 
t2.[11B_3A], 
t2.[12_2A], 
t2.[12_3A], 
t2.[12_4A], 
t2.[11A_3B], 
t2.[11B_3B] 
FROM (SELECT * FROM (SELECT * FROM GrandUnion AS g INNER JOIN org_team_map AS s ON g.org_team_mapID = s.org_team_mapID)  AS t1 INNER JOIN units_summary AS u ON t1.unitID = u.unitID)  AS t2;

'Workday Role_Mapping_By Role'
SELECT DISTINCT *
FROM (SELECT b.Unit, 
b.[Home Dpt Budget Number], 
b.[Employee First Name], 
b.[Employee Last Name], 
b.EID, 
b.[Supervisory Org], 
IIf(Len([b].[7_4A])>0,"x","") AS I9, 
IIF(Len([b].[4A_2A]&[b].[4B_2A]&[b].[4B_3A]&[b].[4C_3A]&[b].[8A_2B])>0,"x","") AS ABP, 
IIf(Len([b].[5B_3A] & [b].[6B_2A] & [b].[8B_2A] & [b].[11A_3A] & [b].[11B_3A])>0,"x","") AS ACP, 
IIf(Len([b].[10_3A])>0,"x","") AS CP, IIf(Len([b].[6A_3A] & [b].[6B_3A] & [b].[12_4A])>0,"x","") AS CAC, 
IIf(Len([b].[5A_2B] & [b].[5A_4A] & [b].[5B_2B] & [b].[5B_4A] & [b].[6B_1A] & [b].[8A_1A] & [b].[8A_3A] & [b].[8B_1B] & [b].[8B_3A] & [b].[8b_4A] & [b].[8C_3A] & [b].[9A_2A] & [b].[9B_2B] & [b].[10_2A] & [b].[11A_2A] & [b].[11A_4A] & [b].[11B_2A] & [b].[12_2A])>0,"x","") AS HRC, 
IIf(Len([b].[6A_2A] & [b].[8A_2A] & [b].[8C_3A] & [b].[9A_3A] & [b].[12_3A])>0,"x","") AS HRP, 
IIf(Len([b].[1_2C] & [b].[1_3B] & [b].[2_2a])>0,"x","") AS TC 
FROM Workday_Role_Mapping_by_field_in_order_of_scenario AS b 
WHERE (((b.[Employee First Name])<>"Ex: Peter")))  AS m;

'Unique Employee Roles'
SELECT DISTINCT
budget_no,
EID, 
first_name,
last_name,
org_team_corrected, 
org_team_mapID,
unitID 
FROM (
SELECT
budget_no,
EID,
first_name,
last_name,
org_team_corrected,
t1.org_team_mapID,
unit_map_ID
FROM (SELECT 
r.EID,
r.budget_no,
r.first_name,
r.last_name,
r.org_team_mapID,
s.unit_map_id
FROM submittals s,
responses r 
WHERE s.submittalID = r.submittalID) t1,
org_team_map o where o.org_team_mapID = t1.org_team_mapID) t2,
unit_correction_map uc where t2.unit_map_ID = uc.unit_map_ID

'submittal_summary'
SELECT t2.submittalID,
t2.date_submitted,
t2.contact,
t2.date_recorded,
t2.session_no,
t2.file_name,
t2.unit_join, 
t2.unit_name,
t2.unitID,
COUNT(r.responseID) As EIDs
FROM (
SELECT
t1.submittalID,
t1.date_submitted,
t1.contact,
t1.date_recorded,
t1.session_no,
t1.file_name,
t1.unitID,
u.unit_join,
u.unit_name 
FROM (SELECT s.submittalID, 
s.date_submitted, 
s.contact,
s.date_recorded, 
s.session_no, 
s.file_name,
um.unitID
FROM submittals AS s, 
unit_correction_map AS um
WHERE s.unit_map_ID=um.unit_map_ID
AND s.superceded = FALSE) t1,
units_summary u
WHERE t1.unitID = u.unitID) t2,
responses r 
WHERE t2.submittalID = r.submittalID
GROUP BY 
t2.submittalID,
t2.date_submitted,
t2.contact,
t2.date_recorded,
t2.session_no,
t2.file_name,
t2.unit_join,
t2.unit_name,
t2.unitID;

