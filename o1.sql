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

select id from oatime where d1 is null and d2 is null and d3 is null and d4 is null and d5 is null and d6 is null ;
select * from oatime where d1<>'缺勤' and d2 ='缺勤' ;
select * from oatime where d2<>'缺勤' and d1 ='缺勤' ;
select * from oatime where d3<>'缺勤' and d4 ='缺勤';
select * from oatime where d4<>'缺勤' and d3 ='缺勤' ;
select * from oatime where d5<>'缺勤' and d6 ='缺勤' ;
select * from oatime where d6<>'缺勤' and d5 ='缺勤' ;


select * from oatime where d1<>'' and d2 is null ;
select * from oatime where d2<>'' and d1 is null ;
select * from oatime where d3<>'' and d4 is null ;
select * from oatime where d4<>'' and d3 is null ;
select * from oatime where d5<>'' and d6 is null ;
select * from oatime where d6<>'' and d5 is null ;