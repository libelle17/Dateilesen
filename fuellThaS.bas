#if vomprogrammeeingelesen then
CREATE DEFINER=`praxis`@`%` FUNCTION `fuellTha`(`inpid` INT
) RETURNS int(11)
    COMMENT 'Fuelle Therapiearten'
BEGIN
DECLARE rc INTEGER(10); SET @inpid=inpid; DELETE FROM therarten WHERE @inpid=0 OR pat_id IN (@inpid); SET @vzahl = ROW_COUNT(); INSERT INTO therarten(pat_id,zp,mpnr,therart,insart,grund,abspos,aktzeit,stbyte) 
SELECT pid,zp,mpnr,thart,ia,gru,abspos,aktzeit,stbyte FROM (
SELECT pid,zp,mpnr,thart, COALESCE(LAG(thart,1) OVER (PARTITION BY pid ORDER BY zp,MPNr),'') lthart
,ia,gru,abspos,NOW() aktzeit,stbyte FROM ( 
WITH dsort AS ( 
SELECT RANK() OVER (PARTITION BY pid ORDER BY zp,MPNr) rang, i.* FROM ( 

  SELECT a.pat_id pid, aufndat Zp, -1 MPNr, aufndat Bis, 'CSII' Thart, 'Anamnese: Insulinpumpe' Gru,  0 ia, x.absPos, x.StByte, 0 FeldNr FROM namen x LEFT JOIN anamnesebogen a USING (pat_id)
  WHERE insulinpumpe<>0 AND (@inpid='0' OR x.pat_id IN (@inpid))
 UNION 
  SELECT Pat_id pid, zeitpunkt Zp,-21 MPNr,ADDDATE(zeitpunkt,365) bis,'CSII' Thart, feldinh Gru,5 ia,FID absPos,Form_id StByte,feldnr FROM formular x
  WHERE form_abk IN ('rp','lar','prp','plar') AND feld IN ('medikament','txtMedKey','VerordnungsZeile') AND feldinh RLIKE 'reservoir|rapid d link|rap d li|rapid-d li|tenderl|flexl|sure t|paradigm|^veo|[^a]veo|animas|cartridge|t-slim|t:slim|t slim|variosoft|trusteel|autosoft|ypsopump|insigh|omnipod' AND NOT feldinh LIKE '%menveo%'
  AND (@inpid='0' OR x.pat_id IN (@inpid))
 UNION 
  SELECT Pat_id pid, zeitpunkt Zp,-22 MPNr,ADDDATE(zeitpunkt,365) bis,'CSII' Thart, Medikament Gru,5 ia,FID absPos, StByte, CASE rezklkurz WHEN'rp'THEN 1 WHEN'lar'THEN 2 ELSE 0 END feldnr FROM rezepteintraege x
  WHERE medikament RLIKE 'reservoir|rapid d link|rap d li|rapid-d li|tenderl|flexl|sure t|paradigm|^veo|[^a]veo|animas|cartridge|t-slim|t:slim|t slim|trusteel|variosoft|trusteel|autosoft|ypsopump|insigh|omnipod'
  AND (@inpid='0' OR x.pat_id IN (@inpid))
 UNION 
  SELECT Pat_id pid, zeitpunkt Zp,-23 MPNr,zeitpunkt bis,'ICT' Thart, feldinh Gru,4 ia,FID absPos,Form_id StByte, feldnr FROM formular x
  WHERE form_abk IN ('rp','lar','prp','plar') AND feld IN ('medikament','txtMedKey','VerordnungsZeile') AND feldinh RLIKE 'fine|Nad|micro fi|[0-9] {0,1}mm|lantus|tresiba|levemir|basal|protaphan|semglee|abasaglar|semilente' AND NOT feldinh RLIKE 'fine {0,1}touch|Katheter|Paradigm|Mio|Flexl|Tender|d li|link|autosoft|Minimed|Sohlen|Oberarm|Fußbett|Schuh|Wanderh|Lancets fine|easy(-release| set)|quick {0,1}set|mmHg|insight|mylife|inset|insulinset|infusion|polster|szinti|dana|fexo|enadura|infektionsnadeln|port|magnes|medtronic|sicherheitslan|omnipod|orbisoft|orbit soft|nadellanz|schlauch|sure|tamponade|trusteel|varisoft|knoten|fine point|flex link|microlet fine|verkürz|vasofix|truesteel|sterile lanzetten|stahlnad|[68]0 {0,1}cm|alkohol|thin lanc|haut|TESTSTR|variosoft|Nadellänge|^BD Micro Fine Lancetten G 33 200 Stück$'
  AND (@inpid='0' OR x.pat_id IN (@inpid))
 UNION 
  (SELECT pat_id pid,qdm zp,-3 MPNr,qdm bis,'ICT' Thart, MID(NAME,p) Gru,-2 ia,x.absPos,x.StByte,0 FROM (SELECT IF(p1>p2,p1,p2) p, b.* FROM (SELECT INSTR(b.name,'insulin') p1, INSTR(b.name,'spritz') p2, b.* FROM tmbrie b) b) x WHERE name RLIKE '(insulin|spritz).*(plan|schema|tabelle)(?!anford)'
   AND (@inpid='0' OR x.pat_id IN (@inpid))

 UNION 
  SELECT DISTINCT pat_id pid, zeitpunkt zp, -5 MPNr, zeitpunkt + INTERVAL 180 DAY bis, 'ICT' Thar, CONCAT('/Insulinplan ',zeitpunkt) gru, -2 ia, 1 absPos,1 StByte, 0
  FROM eintraege x WHERE ((@inpid='0' OR x.pat_id IN (@inpid)) AND (art='iplan' AND inhalt LIKE 'Insulinplan Stand: %'))
 UNION (
	SELECT i.pat_id pid,i.zeitpunkt zp,-6 MPNr,i.zeitpunkt+INTERVAL 6 MONTH bis
	,CASE WHEN ian=3 OR ian=1 THEN 'ICT' ELSE 'CT' END Thar
	,CONCAT('/',GROUP_CONCAT(inh))gru
	,ian ia,1 absPos,1 StByte,0
	FROM(
	SELECT SUM(COALESCE(insart,0))OVER(PARTITION BY eid)ian,i.* FROM(
	SELECT x.pat_id,zeitpunkt,TRIM(REGEXP_REPLACE(inhalt,'^.*Esss?ensinsulin: *([^B]*)Basalinsulin:.*$','\\1'))inh,x.id eid,ma.insart
	FROM eintraege x
	LEFT JOIN medarten ma ON REGEXP_REPLACE(inhalt,'^.*Esss?ensinsulin: *([^ ]*).*$','\\1')=ma.medikament
	WHERE art='iplan' AND (@inpid='0' OR x.pat_id IN (@inpid)) AND inhalt LIKE 'Insulinplan für%'
	GROUP BY x.id,insart
	UNION all
	SELECT x.pat_id,zeitpunkt,TRIM(REGEXP_REPLACE(inhalt,'^.*Basalinsulin: *([^I]*)Insulin [[]iE.*$','\\1')),x.id,ma.insart
	FROM eintraege x
	LEFT JOIN medarten ma ON REGEXP_REPLACE(inhalt,'^.*Basalinsulin: *([^ ]*).*$','\\1')=ma.medikament
	WHERE art='iplan' AND (@inpid='0' OR x.pat_id IN (@inpid)) AND inhalt LIKE 'Insulinplan für%'
	GROUP BY x.id,insart
	)i 
	)i GROUP BY eid )
 UNION 
  SELECT DISTINCT pat_id pid, zeitpunkt zp, -7 MPNr, zeitpunkt + INTERVAL 180 DAY bis, 'ICT' Thar, CONCAT('/Insulinplan ',zeitpunkt) gru, -2 ia, 1 absPos,1 StByte, 0
  FROM eintraege x FORCE INDEX (patid_inhalt) WHERE (@inpid='0' OR x.pat_id IN (@inpid)) AND inhalt RLIKE '^Insulinplan(?! für)'
  GROUP BY pid, zp) 
 UNION 

  SELECT Pid,Zp,MPNr,Zp bis,Thart
  ,REPLACE(REPLACE(REPLACE(gru,'    ',' '),'   ',' '),'  ',' ') gru
  ,ia,abspos,stbyte,feldnr FROM ( 
   SELECT Pid,Zp,MPNr,Zp bis 
       ,CASE WHEN pu THEN 'CSII' 
       WHEN obmzi OR iz>=3 THEN 
       CASE 
        WHEN glp THEN 'GLP1ICT' 
        ELSE 'ICT' 
       END 
      WHEN iz=2 THEN 
       CASE 
        WHEN glp THEN 'GLP1Ins' 
        ELSE 'CT' 
       END 
      WHEN iz=1 THEN 
       CASE 
        WHEN glp THEN 'GLP1Ins' 
        WHEN oboad THEN 'Komb' 
        ELSE 'CT' 
       END 
      ELSE 
       CASE 
        WHEN glp THEN 'GLP1' 
        WHEN oboad THEN 'OAD' 
        ELSE 'Diät' 
       END 
     END Thart 
   ,GROUP_CONCAT(DISTINCT gru SEPARATOR '') gru,MAX(ia) ia
  
	,abspos,stbyte
	,GROUP_CONCAT(feldnr) feldnr 
	
   FROM ( 
    SELECT  Pid,Zp,MPNr,Med
	   ,SUM(Pu) OVER(PARTITION BY pid,zp) pu
	   ,SUM(oboad) OVER(PARTITION BY pid,zp) oboad
	   ,SUM(IF(obin,ezm,0)) OVER(PARTITION BY pid,zp) iz
	 	,SUM(ezm AND ia=1) OVER(PARTITION BY pid,zp) obmzi 
      ,SUM(IF(eztm,glp,0)) OVER(PARTITION BY pid,zp) glp
	   ,MAX(IF(ezm,obin,0)) OVER(PARTITION BY pid,zp) ins
	   ,MAX(IF(ezm,ia,0)) OVER(PARTITION BY pid,zp) ia 
      ,IF(pu||if(eztm,glp,0)||oboad||(ezm&&ia=1)||if(obin,ezm,0),CONCAT('/',Med,' '),'') gru 
      ,IF(pu||if(eztm,glp,0)||oboad||(ezm&&ia=1)||if(obin,ezm,0),Feldnr,NULL) FeldNr
		,absPos
		,StByte 
    FROM ( 
     SELECT pid,Zp,MPNr,Med,pu,ez,wglp,ohneE,IF(ez,oad,0)oboad,glp,obin,ia,FeldNr,absPos,StByte 
      ,IF(ez,ez,ohnee) ezm 
      ,IF(ez,ez,wglp&&ohneE) eztm 
     FROM ( 
      SELECT x.Pat_id pid,x.Zeitpunkt Zp,x.MPNr MPNr,x.Medikament Med,ma.puzu<>0 pu 
       ,MAX((COALESCE(x.mo,'')<>'')+(COALESCE(x.mi,'')<>'')+(COALESCE(x.nm,'')<>'')+(COALESCE(x.ab,'')<>'')+(COALESCE(x.zn,'')<>'')+IF(glp1<>0 AND x.Medanfang RLIKE 'OZEMPIC|TRULICITY|bydureon|Byetta|victoza|Liraglutid|mounjaro',1,0)) ez 
       ,glp1<>0 AND x.MedAnfang RLIKE 'OZEMPIC|TRULICITY|bydureon|Byetta|victoza|Liraglutid|mounjaro' wglp 
       ,pzn<>0 AND concat(x.Bemerkung,' ',x.Grund) NOT RLIKE 'Pause|abgesetzt|beendet|zur Zeit nicht' ohneE 
       ,ma.glib<>0 OR ma.metf<>0 OR ma.gluci<>0 OR ma.shglin<>0 OR ma.dpp4<>0 OR ma.sglt2<>0 OR ma.sonstad<>0 oad 
       ,glp1<>0 glp 
       ,ma.ins<>0 OR ma.anal<>0 obin 
       ,IF(insart='' OR ISNULL(insart),0,insart) ia 
       ,x.FeldNr,x.absPos,x.StByte 
      FROM wmedplan x LEFT JOIN medarten ma ON x.medanfang= ma.medikament 
      WHERE (@inpid='0' OR x.pat_id IN (@inpid))
      GROUP BY x.pat_id,x.zeitpunkt,mpnr,ma.id 
     ) i 
    ) i 
   ) i GROUP BY pid,zp 
  ) i 
 UNION 
  SELECT pat_id pid, insdat zp, -4 MPNr, insdat+INTERVAL 1 DAY bis, IF(Insanw>1,'CSII',NULL) Thart, CONCAT('Insanw: ',insanw) Gru,4 ia, 0 absPos,StByte, -1 feldnr FROM namen x 
  WHERE (@inpid='0' OR x.pat_id IN (@inpid))
 ) i 
)  
SELECT d.*
, COALESCE(ldso.thart,'') lathart 
, RANK() over (PARTITION BY pid ORDER BY zp,mpnr) thrang
, COALESCE(LAG(d.zp,1) OVER (PARTITION BY pid ORDER BY zp,MPNr),'1900-01-01') lzp
, COALESCE(LAG(d.ia,1) OVER (PARTITION BY pid ORDER BY zp,MPNr),'-10') lia 
FROM dsort d LEFT JOIN dsort ldso 
 ON ldso.pid=d.pid AND ldso.rang=(SELECT MAX(rang) FROM dsort WHERE pid=d.pid AND rang<d.rang AND thart<>d.thart)
 WHERE NOT EXISTS (SELECT 1 FROM dsort WHERE pid=d.pid AND rang<d.rang AND (zp<d.zp AND bis>=d.zp)) 
) i 
WHERE FALSE OR thart<>lathart AND NOT (
  thart='Diät' AND lathart IN ('CSII','ICT','GLP1ICT') AND NOT EXISTS (SELECT 1 FROM sws WHERE pat_id=i.pid AND voret BETWEEN i.lzp AND i.zp))
  AND NOT (lia=-2 AND ia=2 AND NOT (thart<>lathart AND thart IN ('GLP1','GLP1Ins','GLP1ICT')) AND zp BETWEEN lzp AND ADDDATE(lzp,92)) 
  AND NOT (lathart='CSII' AND ia<>4)
GROUP BY pid,zp,thart
ORDER BY pid,zp,MPNr
) i WHERE thart<>lthart
; 
SET rc=ROW_COUNT(); 
UPDATE anamnesebogen x SET ther1=(SELECT therart FROM therarten WHERE pat_id=x.pat_id ORDER BY zp, mpnr LIMIT 1), therakt =(SELECT therart FROM therarten WHERE pat_id=x.pat_id ORDER BY zp DESC, mpnr DESC LIMIT 1) WHERE (@inpid='0' OR x.pat_id IN (@inpid)); 
RETURN rc; 
END
#endif ' vomprogrammeeingelesen then
