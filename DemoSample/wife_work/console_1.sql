-- 整理总排名，根据赋分降序排列,文理科自己切换，执行
-- 查询物理类所有学生，按照等级总分降序排列
-- 自己改写代码，区分文理，修改 where 类别 not in ('历政地')
select z.姓名,
       z.班级,
       z.类别,
       z.语文分数,
       z.数学分数,
       z.英语分数,
       z.物理分数,
       z.历史分数,
       z.化学分数,
       z.生物分数,
       z.政治分数,
       z.地理分数,
       z.总分,
       z.等级总分,
       RANK() OVER (PARTITION BY z.年级, z.班级 ORDER BY z.总分 DESC) AS 班名,
       RANK() OVER (ORDER BY 等级总分 DESC)                           AS 类别名次,
       s.名次                                                         as 上次考试名次,
       s.名次 - RANK() OVER (ORDER BY 等级总分 DESC)                  as 名次差
from 总成绩单 z
         left join "上次排名" s on z.姓名 = s.姓名 and z.年级 = s.年级 and z.班级 = s.班级

--where 类别 not in ('历政地')
where 类别 not in ('历政地')
order by 等级总分 desc;


-- 生成扣人后的总名单,文理科自己切换，执行
select z.姓名,
       z.班级,
       z.类别,
       z.语文分数,
       z.数学分数,
       z.英语分数,
       z.物理分数,
       z.历史分数,
       z.化学分数,
       z.生物分数,
       z.政治分数,
       z.地理分数,
       z.总分,
       z.等级总分,
       RANK() OVER (PARTITION BY z.年级, z.班级 ORDER BY z.总分 DESC) AS 班名,
       RANK() OVER (ORDER BY 等级总分 DESC)                           AS 类别名次,
       s.名次                                                         as 上次考试名次,
       s.名次 - RANK() OVER (ORDER BY 等级总分 DESC)                  as 名次差
from 总成绩单 z
         left join "上次排名" s on z.姓名 = s.姓名 and z.年级 = s.年级 and z.班级 = s.班级

--where 类别 not in ('历政地')
where 类别 not in ('历政地')
  and (z.姓名, z.班级, z.年级) not in (select 扣人名单.姓名, 扣人名单.班级, 扣人名单.年级 from 扣人名单)
order by 等级总分 desc;


-- 每个班各科成绩的及格率，低分率，平均分，参评人数
--各科成绩统计表
SELECT 年级,
       班级,

       ROUND(CASE
                 WHEN SUM(语文分数) = 0 THEN 0
                 ELSE
                     count(case when 语文分数 != 0 then 1 end)
                 END, 2) AS 语文参评人数,

       ROUND(CASE
                 WHEN SUM(语文分数) = 0 THEN 0
                 ELSE
                     AVG(CASE WHEN 语文分数 != 0 THEN 语文分数 END)
                 END, 2) AS 语文平均分,

       ROUND(CASE
                 WHEN SUM(语文分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 语文分数 >= 90 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 语文分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 语文及格率,

       ROUND(CASE
                 WHEN SUM(语文分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 语文分数 <= 60 and 语文分数 != 0 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 语文分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 语文低分率,

       ROUND(CASE
                 WHEN SUM(数学分数) = 0 THEN 0
                 ELSE
                     count(case when 数学分数 != 0 then 1 end)
                 END, 2) AS 数学参评人数,

       ROUND(CASE
                 WHEN SUM(数学分数) = 0 THEN 0
                 ELSE
                     AVG(CASE WHEN 数学分数 != 0 THEN 数学分数 END)
                 END, 2) AS 数学平均分,

       ROUND(CASE
                 WHEN SUM(数学分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 数学分数 >= 90 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 数学分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 数学及格率,

       ROUND(CASE
                 WHEN SUM(数学分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 数学分数 <= 60 and 数学分数 != 0 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 数学分数 != 0  THEN 1 ELSE 0 END)
                 END, 2) AS 数学低分率,
---------------------------
       ROUND(CASE
                 WHEN SUM(英语分数) = 0 THEN 0
                 ELSE
                     count(case when 英语分数 != 0 then 1 end)
                 END, 2) AS 英语参评人数,

       ROUND(CASE
                 WHEN SUM(英语分数) = 0 THEN 0
                 ELSE
                     AVG(CASE WHEN 英语分数 != 0 THEN 英语分数 END)
                 END, 2) AS 英语平均分,

       ROUND(CASE
                 WHEN SUM(英语分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 英语分数 >= 90 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 英语分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 英语及格率,

       ROUND(CASE
                 WHEN SUM(英语分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 英语分数 <= 60 and 总成绩单.英语分数 != 0 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 英语分数 != 0  THEN 1 ELSE 0 END)
                 END, 2) AS 英语低分率,


       ROUND(CASE
                 WHEN SUM(物理分数) = 0 THEN 0
                 ELSE
                     count(case when 物理分数 != 0 then 1 end)
                 END, 2) AS 物理参评人数,

       ROUND(CASE
                 WHEN SUM(物理分数) = 0 THEN 0
                 ELSE
                     AVG(CASE WHEN 物理分数 != 0 THEN 物理分数 END)
                 END, 2) AS 物理平均分,

       ROUND(CASE
                 WHEN SUM(物理分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 物理分数 >= 60 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 物理分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 物理及格率,

       ROUND(CASE
                 WHEN SUM(物理分数) = 0 THEN 0
                 ELSE
                        SUM(CASE WHEN 物理分数 <= 40 and 物理分数 != 0 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 物理分数 != 0  THEN 1 ELSE 0 END)
                 END, 2) AS 物理低分率,


       ROUND(CASE
                 WHEN SUM(历史分数) = 0 THEN 0
                 ELSE
                     count(case when 历史分数 != 0 then 1 end)
                 END, 2) AS 历史参评人数,

       ROUND(CASE
                 WHEN SUM(历史分数) = 0 THEN 0
                 ELSE
                     AVG(CASE WHEN 历史分数 != 0 THEN 历史分数 END)
                 END, 2) AS 历史平均分,

       ROUND(CASE
                 WHEN SUM(历史分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 历史分数 >= 60 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 历史分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 历史及格率,

       ROUND(CASE
                 WHEN SUM(历史分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 历史分数 <= 40 and 历史分数 != 0 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 历史分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 历史低分率,


       ROUND(CASE
                 WHEN SUM(化学分数) = 0 THEN 0
                 ELSE
                     count(case when 化学分数 != 0 then 1 end)
                 END, 2) AS 化学参评人数,

       ROUND(CASE
                 WHEN SUM(化学分数) = 0 THEN 0
                 ELSE
                     AVG(CASE WHEN 化学分数 != 0 THEN 化学分数 END)
                 END, 2) AS 化学平均分,

       ROUND(CASE
                 WHEN SUM(化学分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 化学分数 >= 60 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 化学分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 化学及格率,

       ROUND(CASE
                 WHEN SUM(化学分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 化学分数 <= 40  and 化学分数 != 0 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 化学分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 化学低分率,


       ROUND(CASE
                 WHEN SUM(生物分数) = 0 THEN 0
                 ELSE
                     count(case when 生物分数 != 0 then 1 end)
                 END, 2) AS 生物参评人数,

       ROUND(CASE
                 WHEN SUM(生物分数) = 0 THEN 0
                 ELSE
                     AVG(CASE WHEN 生物分数 != 0 THEN 生物分数 END)
                 END, 2) AS 生物平均分,

       ROUND(CASE
                 WHEN SUM(生物分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 生物分数 >= 60 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 生物分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 生物及格率,

       ROUND(CASE
                 WHEN SUM(生物分数) = 0 THEN 0
                 ELSE
                        SUM(CASE WHEN 生物分数 <= 40  and 生物分数 != 0 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 生物分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 生物低分率,

       ROUND(CASE
                 WHEN SUM(政治分数) = 0 THEN 0
                 ELSE
                     count(case when 政治分数 != 0 then 1 end)
                 END, 2) AS 政治参评人数,

       ROUND(CASE
                 WHEN SUM(政治分数) = 0 THEN 0
                 ELSE
                     AVG(CASE WHEN 政治分数 != 0 THEN 政治分数 END)
                 END, 2) AS 政治平均分,

       ROUND(CASE
                 WHEN SUM(政治分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 政治分数 >= 60 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 政治分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 政治及格率,

       ROUND(CASE
                 WHEN SUM(政治分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 政治分数 <= 40 and 政治分数 != 0 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 政治分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 政治低分率,

       ROUND(CASE
                 WHEN SUM(地理分数) = 0 THEN 0
                 ELSE
                     count(case when 地理分数 != 0 then 1 end)
                 END, 2) AS 地理参评人数,

       ROUND(CASE
                 WHEN SUM(地理分数) = 0 THEN 0
                 ELSE
                     AVG(CASE WHEN 地理分数 != 0 THEN 地理分数 END)
                 END, 2) AS 地理平均分,

       ROUND(CASE
                 WHEN SUM(地理分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 地理分数 >= 60 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 地理分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 地理及格率,

       ROUND(CASE
                 WHEN SUM(地理分数) = 0 THEN 0
                 ELSE
                         SUM(CASE WHEN 地理分数 <= 40 and 地理分数 != 0 THEN 1 ELSE 0 END) /
                         SUM(CASE WHEN 地理分数 != 0 THEN 1 ELSE 0 END)
                 END, 2) AS 地理低分率


FROM 总成绩单
where (姓名, 班级, 年级) not in (select 扣人名单.姓名, 扣人名单.班级, 扣人名单.年级 from 扣人名单)
GROUP BY 年级, 班级
order by 班级;

--获取每个班中不同类别的总分平均分
SELECT 年级,
       班级,
       类别,
       ROUND(AVG(总分), 2) AS 总分平均分
FROM 总成绩单
where (姓名, 班级, 年级) not in (select 扣人名单.姓名, 扣人名单.班级, 扣人名单.年级 from 扣人名单)
group by 年级, 班级, 类别;



--班级总成绩统计
SELECT 年级,班级, 类别, COUNT(*) AS 参评人数
FROM 总成绩单
WHERE
    (姓名, 班级, 年级) not in (select 扣人名单.姓名, 扣人名单.班级, 扣人名单.年级 from 扣人名单)
and
    (语文分数 > 0 AND 英语分数 > 0 AND 数学分数 > 0)
    AND (
        -- 若有类别，则自行确认类别科目，
        (类别 IN ('物化政') AND 物理分数 > 0 AND 化学分数 > 0 AND 政治分数 > 0)
        OR (类别 IN ('物化地') AND 物理分数 > 0 AND 化学分数 > 0 AND 地理分数 > 0)
        OR (类别 IN ('物生政') AND 物理分数 > 0 AND 政治分数 > 0 AND 生物分数 > 0)
        OR (类别 IN ('物政地') AND 物理分数 > 0 AND 政治分数 > 0 AND 地理分数 > 0)
        OR (类别 IN ('物生地') AND 物理分数 > 0 AND 生物分数 > 0 AND 地理分数 > 0)
        OR (类别 IN ('政历地') AND 政治分数 > 0 AND 历史分数 > 0 AND 地理分数 > 0)

        -- 若没有类别，则判断全部科目成绩
        OR (类别 IS NULL AND 物理分数 > 0 AND 化学分数 > 0 AND 生物分数 > 0 AND 政治分数 > 0 AND 历史分数 > 0 AND 地理分数 > 0)
    )
GROUP BY 年级,班级, 类别
order by 年级,班级;


