SELECT *
FROM oatime
WHERE d5 NOT LIKE '缺%' AND d6 LIKE '缺%';

SELECT *
FROM oatime
WHERE name = '张世玉';

SELECT *
FROM oatime
WHERE d4 LIKE '缺%' AND d5 NOT LIKE '缺%' AND d6 NOT LIKE '缺%';


UPDATE oatime1
SET d4 = "18:00", d5 = "18:30"
WHERE id IN (SELECT id
             FROM oatime
             WHERE d4 LIKE '缺%' AND d5 NOT LIKE '缺%' AND d6 NOT LIKE '缺%');

SELECT *
FROM oatime
WHERE d1 LIKE '%60' OR d2 LIKE '%60' OR d3 LIKE '%60' OR d4 LIKE '%60' OR d5 LIKE '%60' OR d6 LIKE '%60';

