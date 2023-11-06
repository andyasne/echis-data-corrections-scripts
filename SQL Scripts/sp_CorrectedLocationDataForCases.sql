/*
   getFamilyCorrections - Retrieve family corrections based on specified criteria
   
   This stored procedure retrieves family corrections from the 'family' table. It returns a list of unique case identifiers, owner IDs, and case types for cases where the owner ID matches 'from_location_id' and the village name matches either 'from_village_name1,' 'from_village_name2,' or 'from_village_name3.'
   
   Parameters:
   - to_location_id: The owner ID to be used in the results.
   - from_location_id: The owner ID to match in the 'family' table.
   - from_village_name1: The first village name to match.
   - from_village_name2: The second village name to match.
   - from_village_name3: The third village name to match.
   
   Returns:
   - A table containing case identifiers, owner IDs, and case types.
   
   Usage example:
   SELECT * FROM getFamilyCorrections('to_location_id_value', 'from_location_id_value', 'village_name1_value', 'village_name2_value', 'village_name3_value');
*/

CREATE OR REPLACE FUNCTION getFamilyCorrections(to_location_id text, from_location_id text, from_village_name1 text, from_village_name2 text, from_village_name3 text)
RETURNS TABLE(caseid text, owner_id text, case_type text) AS
$$
BEGIN
  RETURN QUERY
  SELECT DISTINCT f.caseid, to_location_id AS owner_id, f.case_type
  FROM family f
  WHERE f.owner_id = from_location_id
  AND (f.village_name = from_village_name1 OR f.village_name = from_village_name2 OR f.village_name = from_village_name3);
END;
$$
LANGUAGE 'plpgsql';




//TODO: do stored procedure for the following 
SELECT c.caseid ,to_location_id as owner_id, c.case_type FROM 
(select distinct  c.caseid , c.case_type ,c."indices.family"  from Client ) as c 
inner join
 ( select caseid from 
getFamilyCorrections(to_location_id,from_location_id, village_name1_value, village_name2_value, village_name3_value)) as p ON c."indices.family" = p.caseid ;

