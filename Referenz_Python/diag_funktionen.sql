DELIMITER $$
DROP FUNCTION IF EXISTS diag_clean$$
CREATE FUNCTION diag_clean(t LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN RETURN TRIM(REGEXP_REPLACE(COALESCE(t,''), '\\s+', ' ')); END$$
DROP FUNCTION IF EXISTS diag_suffix$$
CREATE FUNCTION diag_suffix(t LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN
  SET t = REGEXP_REPLACE(COALESCE(t,''), '\\s*\\[.*$', '');
  SET t = REGEXP_REPLACE(t, '\\s*#+\\s*$', '');
  RETURN REGEXP_REPLACE(t, '[\\s,;:.\\-]+$', '');
END$$
DROP FUNCTION IF EXISTS diag_strip$$
CREATE FUNCTION diag_strip(t LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN
  DECLARE prev LONGTEXT DEFAULT NULL; DECLARE n INT DEFAULT 0;
  SET t = COALESCE(t,'');
  WHILE (prev IS NULL OR prev <> t) AND n < 12 DO
    SET prev = t; SET n = n + 1;
    SET t = REGEXP_REPLACE(t, '(?i)^\\s*(?:[\\[(]\\s*(?:wohl\\s+)?(?:Z\\.?\\s?n\\.?|V\\.?\\s?a\\.?|[ZVGA])\\s*\\??\\s*[\\])]|[\\[(][A-Z]\\d\\d[\\dA-Z.\\-+*!]*[\\])]|[A-Z]\\d\\d\\.\\d{1,2}[\\dA-Z]?\\s|Z\\.?\\s?n\\.?(?=\\s)|Zn\\.|Zustand\\s+nach|V\\.\\s?a\\.|Verdacht\\s+auf|Verdacht|gesichert|ausgeschl\\.?|Ausschluss|ausgeschlossen|DD\\b\\.?)\\s*[:,\\-]?\\s*', '');
  END WHILE;
  RETURN t;
END$$
DROP FUNCTION IF EXISTS diag_s$$
CREATE FUNCTION diag_s(t LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN RETURN diag_clean(diag_suffix(diag_strip(t))); END$$
DROP FUNCTION IF EXISTS diag_nodate$$
CREATE FUNCTION diag_nodate(s LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN RETURN REGEXP_REPLACE(s, '(?i)\\s*(?:(?:seit|bis|ab|vor|am|ED)\\s+)?(?:\\d{1,2}\\.\\d{1,2}\\.(?:\\d{2,4})?|\\d{1,2}/\\d{2,4}|(?:19|20)\\d\\d)\\b\\.?', ''); END$$
DROP FUNCTION IF EXISTS diag_dates$$
CREATE FUNCTION diag_dates(s LONGTEXT) RETURNS VARCHAR(500) DETERMINISTIC NO SQL
BEGIN
  DECLARE x LONGTEXT; DECLARE m VARCHAR(30); DECLARE res VARCHAR(500) DEFAULT ''; DECLARE n INT DEFAULT 0;
  SET x = COALESCE(s,'');
  WHILE n < 20 DO
    SET m = REGEXP_SUBSTR(x, '\\d{1,2}\\.\\d{1,2}\\.(?:\\d{2,4})?|\\d{1,2}/\\d{2,4}|(?:19|20)\\d\\d');
    IF m IS NULL OR m = '' THEN SET n = 20; ELSE
      SET res = CONCAT(res, m, ' '); SET x = INSERT(x, LOCATE(m, x), CHAR_LENGTH(m), ' '); SET n = n + 1;
    END IF;
  END WHILE;
  RETURN TRIM(res);
END$$
DROP FUNCTION IF EXISTS diag_noside$$
CREATE FUNCTION diag_noside(t LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN RETURN diag_clean(REGEXP_REPLACE(t, '(?i)(?<![\\wäöüß])(links|li\\.?|rechts|re\\.?|beidseits|beidseitig[\\wäöüß]*|bds\\.?)(?![\\wäöüß])', ' ')); END$$
DROP FUNCTION IF EXISTS diag_sides$$
CREATE FUNCTION diag_sides(t LONGTEXT, seite VARCHAR(1)) RETURNS VARCHAR(3) DETERMINISTIC NO SQL
BEGIN
  DECLARE l INT DEFAULT 0; DECLARE r INT DEFAULT 0;
  IF t REGEXP '(?i)(?<![\\wäöüß])(links|li\\.?)(?![\\wäöüß])' THEN SET l = 1; END IF;
  IF t REGEXP '(?i)(?<![\\wäöüß])(rechts|re\\.?)(?![\\wäöüß])' THEN SET r = 1; END IF;
  IF t REGEXP '(?i)(?<![\\wäöüß])(beidseits|beidseitig[\\wäöüß]*|bds\\.?)(?![\\wäöüß])' THEN SET l = 1; SET r = 1; END IF;
  IF seite = 'R' THEN SET r = 1; ELSEIF seite = 'L' THEN SET l = 1; ELSEIF seite = 'B' THEN SET l = 1; SET r = 1; END IF;
  RETURN CONCAT(IF(l=1,'L',''), IF(r=1,'R',''));
END$$
DROP FUNCTION IF EXISTS diag_b0$$
CREATE FUNCTION diag_b0(t LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN RETURN LOWER(diag_clean(diag_suffix(diag_noside(diag_nodate(diag_s(t)))))); END$$
DROP FUNCTION IF EXISTS diag_b1$$
CREATE FUNCTION diag_b1(t LONGTEXT) RETURNS LONGTEXT DETERMINISTIC READS SQL DATA
BEGIN
  DECLARE done INT DEFAULT 0; DECLARE m VARCHAR(2000); DECLARE c VARCHAR(80);
  DECLARE cur CURSOR FOR SELECT Muster, Kanonisch FROM diagnormal ORDER BY Reihenfolge;
  DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = 1;
  SET t = LOWER(REPLACE(REPLACE(diag_noside(diag_nodate(diag_s(t))), 'ß', 'ss'), '-', ' '));
  OPEN cur;
  lp: LOOP
    FETCH cur INTO m, c;
    IF done = 1 THEN LEAVE lp; END IF;
    SET t = REGEXP_REPLACE(t, CONCAT('(?i)', m), c);
  END LOOP;
  CLOSE cur;
  SET t = REGEXP_REPLACE(t, '\\s+,', ',');
  SET t = REGEXP_REPLACE(t, '(,\\s*)+', ', ');
  RETURN diag_clean(diag_suffix(t));
END$$
DROP FUNCTION IF EXISTS diag_core$$
CREATE FUNCTION diag_core(b1 LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN RETURN diag_clean(diag_suffix(REGEXP_REPLACE(b1, '\\s*[,(].*$', ''))); END$$
DROP FUNCTION IF EXISTS diag_rest$$
CREATE FUNCTION diag_rest(b1 LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN RETURN TRIM(REGEXP_REPLACE(SUBSTRING(b1, CHAR_LENGTH(diag_core(b1)) + 1), '[^\\wäöüß]+', ' ')); END$$
DROP FUNCTION IF EXISTS diag_tokens$$
CREATE FUNCTION diag_tokens(t LONGTEXT) RETURNS LONGTEXT DETERMINISTIC NO SQL
BEGIN RETURN TRIM(REGEXP_REPLACE(LOWER(COALESCE(t,'')), '[^\\wäöüß]+', ' ')); END$$
DELIMITER ;

DELIMITER $$
DROP FUNCTION IF EXISTS diag_norm1$$
CREATE FUNCTION diag_norm1(t2 LONGTEXT) RETURNS LONGTEXT DETERMINISTIC READS SQL DATA
BEGIN
  DECLARE done INT DEFAULT 0; DECLARE m VARCHAR(2000); DECLARE c VARCHAR(80); DECLARE t LONGTEXT;
  DECLARE cur CURSOR FOR SELECT Muster, Kanonisch FROM diagnormal ORDER BY Reihenfolge;
  DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = 1;
  SET t = LOWER(REPLACE(REPLACE(t2, 'ß', 'ss'), '-', ' '));
  OPEN cur;
  lp: LOOP
    FETCH cur INTO m, c;
    IF done = 1 THEN LEAVE lp; END IF;
    SET t = REGEXP_REPLACE(t, CONCAT('(?i)', m), c);
  END LOOP;
  CLOSE cur;
  SET t = REGEXP_REPLACE(t, '\\s+,', ',');
  SET t = REGEXP_REPLACE(t, '(,\\s*)+', ', ');
  RETURN diag_clean(diag_suffix(t));
END$$
DROP FUNCTION IF EXISTS diag_info$$
CREATE FUNCTION diag_info(t LONGTEXT, seite VARCHAR(1), attr LONGTEXT) RETURNS LONGTEXT DETERMINISTIC READS SQL DATA
BEGIN
  -- Felder getrennt durch CHAR(31): s, dates, b0, b1, core, rest, sides, erl, tnorm
  DECLARE s LONGTEXT; DECLARE t1 LONGTEXT; DECLARE t2 LONGTEXT; DECLARE b1 LONGTEXT; DECLARE core LONGTEXT; DECLARE sep CHAR(1) DEFAULT CHAR(31);
  SET s = diag_s(t);
  SET t1 = diag_nodate(s);
  SET t2 = diag_noside(t1);
  SET b1 = diag_norm1(t2);
  SET core = diag_core(b1);
  RETURN CONCAT_WS(sep, s, diag_dates(s), LOWER(diag_clean(diag_suffix(t2))), b1, core,
                   TRIM(REGEXP_REPLACE(SUBSTRING(b1, CHAR_LENGTH(core) + 1), '[^\\wäöüß]+', ' ')), diag_sides(t1, seite), diag_tokens(attr), LOWER(diag_clean(diag_suffix(t1))));
END$$
DELIMITER ;

DELIMITER $$
DROP FUNCTION IF EXISTS diag_jahr$$
CREATE FUNCTION diag_jahr(t LONGTEXT) RETURNS INT DETERMINISTIC NO SQL
BEGIN
  -- hoechstes Jahr im Text: vierstellig (1900-2099) oder als Monat/Jahr (8/16); 0 = keins
  DECLARE x LONGTEXT; DECLARE m VARCHAR(10); DECLARE best INT DEFAULT 0; DECLARE n INT DEFAULT 0; DECLARE y INT;
  SET x = COALESCE(t, '');
  WHILE n < 20 DO
    SET m = REGEXP_SUBSTR(x, '(?<!\\d)(?:19|20)\\d\\d(?!\\d)');
    IF m IS NULL OR m = '' THEN SET n = 20;
    ELSE
      SET y = CAST(m AS UNSIGNED); IF y > best THEN SET best = y; END IF;
      SET x = INSERT(x, LOCATE(m, x), 4, ' '); SET n = n + 1;
    END IF;
  END WHILE;
  SET x = COALESCE(t, ''); SET n = 0;
  WHILE n < 20 DO
    SET m = REGEXP_SUBSTR(x, '(?<!\\d)(?:0?[1-9]|1[0-2])/\\d\\d(?!\\d)');
    IF m IS NULL OR m = '' THEN SET n = 20;
    ELSE
      SET y = CAST(SUBSTRING_INDEX(m, '/', -1) AS UNSIGNED); SET y = IF(y <= 30, 2000 + y, 1900 + y);
      IF y > best THEN SET best = y; END IF;
      SET x = INSERT(x, LOCATE(m, x), CHAR_LENGTH(m), ' '); SET n = n + 1;
    END IF;
  END WHILE;
  RETURN best;
END$$
DELIMITER ;
