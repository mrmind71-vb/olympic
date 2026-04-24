ALTER VIEW [dbo].[FILE3_11RP]
AS
SELECT     '1' AS TYPE, NULL AS DOC_ID, CODE, F_DATE AS DATE, CASE WHEN F_BAL1 > 0 THEN F_BAL1 ELSE 0 END AS SAL, 
                      CASE WHEN F_BAL1 < 0 THEN ABS(F_BAL1) ELSE 0 END AS PAY, '—’Ìœ «Ê·' AS DESCA
FROM         dbo.FILE3_10
WHERE     (F_BAL1 <> 0)
UNION ALL
SELECT     '4' AS Expr1, MAX(dbo.FILE6_20H.DOC_NO) AS Expr3, dbo.FILE6_20H.code, dbo.FILE6_20H.date, SUM(dbo.FILE6_20.TOTAL) 
                      + dbo.FILE6_20H.tax - dbo.FILE6_20H.discount AS SAL, 0 AS PAY, '
                       „»Ì⁄«  ' AS Expr2
FROM         dbo.FILE6_20H INNER JOIN
                      dbo.FILE6_20 ON dbo.FILE6_20H.DOC_NO = dbo.FILE6_20.DOC_NO
GROUP BY dbo.FILE6_20H.code, dbo.FILE6_20H.date, dbo.FILE6_20H.tax, dbo.FILE6_20H.discount
UNION ALL
SELECT     '10 ' AS Expr1, DOC_NO, code, date, 0 AS SAL, CASH AS PAY, ' „ﬁ»Ê÷«  ‰ﬁœÌ… ' AS Expr2
FROM         dbo.FILE6_20H AS FILE6_20H_1
WHERE     (NOT (BOX IS NULL))
GROUP BY DOC_NO, code, date, tax, discount, CASH
UNION ALL
SELECT     '12' AS Expr1, DOC_NO, code, date, 0 AS SAL, VISA AS PAY, '„ﬁ»Ê÷«  ›Ì“«' AS Expr2
FROM         dbo.FILE6_20H AS FILE6_20H_2
WHERE     (VISA <> 0)
UNION ALL
SELECT     '11' AS Expr1, DOC_NO, code, date, CASH AS SAL, 0 AS PAY, '„œ›Ê⁄«  ‰ﬁœÌ… „—œÊœ „»Ì⁄« ' AS Expr2
FROM         dbo.FILE6_10H
WHERE     (NOT (BOX IS NULL))
UNION ALL
SELECT     '5' AS Expr1, MAX(FILE6_10H_1.DOC_NO) AS Expr3, FILE6_10H_1.code, FILE6_10H_1.date, 0 AS SAL, SUM(dbo.FILE6_10.TOTAL) 
                      + FILE6_10H_1.tax - FILE6_10H_1.discount AS PAY, ' „—œÊœ „»Ì⁄« ' AS Expr2
FROM         dbo.FILE6_10H AS FILE6_10H_1 INNER JOIN
                      dbo.FILE6_10 ON FILE6_10H_1.DOC_NO = dbo.FILE6_10.DOC_NO
GROUP BY FILE6_10H_1.code, FILE6_10H_1.date, FILE6_10H_1.tax, FILE6_10H_1.discount
UNION ALL
SELECT     '7' AS Expr1, dbo.FILE8_10H.DOC_NO, dbo.FILE8_10.CODE, dbo.FILE8_10H.DATE, 0 AS SAL, SUM(dbo.FILE8_10.VALUE) AS PAY, 
                      dbo.FILE8_10.desca
FROM         dbo.FILE8_10H INNER JOIN
                      dbo.FILE8_10 ON dbo.FILE8_10H.DOC_NO = dbo.FILE8_10.DOC_NO
GROUP BY dbo.FILE8_10H.DOC_NO, dbo.FILE8_10.CODE, dbo.FILE8_10H.DATE, dbo.FILE8_10.desca
UNION ALL
SELECT     '8' AS Expr1, dbo.FILE8_30H.DOC_NO, dbo.FILE8_30.CODE, dbo.FILE8_30H.DATE, SUM(dbo.FILE8_30.VALUE) AS SAL, 0 AS PAY, 
                      dbo.FILE8_30.DESCA
FROM         dbo.FILE8_30H INNER JOIN
                      dbo.FILE8_30 ON dbo.FILE8_30H.DOC_NO = dbo.FILE8_30.DOC_NO
GROUP BY dbo.FILE8_30H.DOC_NO, dbo.FILE8_30.CODE, dbo.FILE8_30H.DATE, dbo.FILE8_30.DESCA
UNION ALL
SELECT     'A' AS Expr1, SER_NO, code1, DATE_R, 0 AS SAL, VALUE AS PAY, '«Ê—«ﬁ ﬁ»÷' AS Expr2
FROM         dbo.FILE5_20
WHERE     (NOT (code1 IS NULL)) AND (OLD = 0)
UNION ALL
SELECT     'C' AS Expr1, SER_NO, code1, DATE_3, VALUE AS SAL, 0 AS PAY, '—›÷ «Ê—«ﬁ ﬁ»÷' AS Expr2
FROM         dbo.FILE5_20 AS FILE5_20_1
WHERE     (NOT (code1 IS NULL)) AND (CLOSED = '1')
UNION ALL
SELECT     'E' AS Expr1, SER_NO, transcode2, DATE_R, VALUE AS SAL, 0 AS PAY, ' ŸÂÌ— ‘Ìﬂ' AS Expr2
FROM         dbo.FILE5_20 AS FILE5_20_2
WHERE     (NOT (transcode2 IS NULL))
