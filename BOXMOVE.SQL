ALTER VIEW [dbo].[BOXMOVE]
AS
SELECT     dbo.FILE8_10H.DOC_NO, dbo.FILE8_10.VALUE AS PLUS, 0 AS MINUS, dbo.FILE8_10.BOX, dbo.FILE8_10H.DATE, 1 AS FLAG, '„ﬁ»Ê÷«  „‰ ⁄„·«¡' AS DESCA,
                       dbo.FILE3_10.DESCA AS CODEDESCA
FROM         dbo.FILE8_10H INNER JOIN
                      dbo.FILE8_10 ON dbo.FILE8_10H.DOC_NO = dbo.FILE8_10.DOC_NO INNER JOIN
                      dbo.FILE3_10 ON dbo.FILE8_10.CODE = dbo.FILE3_10.CODE
WHERE     (NOT (dbo.FILE8_10.BOX IS NULL))
UNION ALL
SELECT     dbo.FILE8_20H.DOC_NO, 0 AS PLUS, dbo.FILE8_20.VALUE AS MINUS, dbo.FILE8_20.BOX, dbo.FILE8_20H.DATE, 2 AS FLAG, 
                      '„œ›Ê⁄«  «·Ì „Ê—œÌ‰' AS DESCA, dbo.file4_10.DESCA AS CODEDESCA
FROM         dbo.FILE8_20H INNER JOIN
                      dbo.FILE8_20 ON dbo.FILE8_20H.DOC_NO = dbo.FILE8_20.DOC_NO INNER JOIN
                      dbo.file4_10 ON dbo.FILE8_20.CODE = dbo.file4_10.CODE
WHERE     (NOT (dbo.FILE8_20.BOX IS NULL))
UNION ALL
SELECT     dbo.FILE8_30H.DOC_NO, 0 AS PLUS, dbo.FILE8_30.VALUE AS MINUS, dbo.FILE8_30.BOX, dbo.FILE8_30H.DATE, 3 AS FLAG, 
                      '„œ›Ê⁄«   «·Ì ⁄„·«¡' AS DESCA, FILE3_10_1.DESCA AS CODEDESCA
FROM         dbo.FILE8_30H INNER JOIN
                      dbo.FILE8_30 ON dbo.FILE8_30H.DOC_NO = dbo.FILE8_30.DOC_NO INNER JOIN
                      dbo.FILE3_10 AS FILE3_10_1 ON dbo.FILE8_30.CODE = FILE3_10_1.CODE
WHERE     (NOT (dbo.FILE8_30.BOX IS NULL))
UNION ALL
SELECT     dbo.FILE8_40H.DOC_NO, dbo.FILE8_40.VALUE, 0 AS MINUS, dbo.FILE8_40.BOX, dbo.FILE8_40H.DATE, 4 AS FLAG, '„ﬁ»Ê÷«  „‰ „Ê—œÌ‰' AS DESCA, 
                      file4_10_1.DESCA AS CODEDESCA
FROM         dbo.FILE8_40H INNER JOIN
                      dbo.FILE8_40 ON dbo.FILE8_40H.DOC_NO = dbo.FILE8_40.DOC_NO INNER JOIN
                      dbo.file4_10 AS file4_10_1 ON dbo.FILE8_40.CODE = file4_10_1.CODE
WHERE     (NOT (dbo.FILE8_40.BOX IS NULL))
UNION ALL
SELECT     dbo.FILE8_50H.DOC_NO, 0 AS PLUS, dbo.FILE8_50.VALUE AS MINSU, dbo.FILE8_50.BOX, dbo.FILE8_50H.DATE, 5 AS FLAG, '„’«—Ì›' AS DESCA, 
                      dbo.FILE8_51.DESCA AS Expr1
FROM         dbo.FILE8_50H INNER JOIN
                      dbo.FILE8_50 ON dbo.FILE8_50H.DOC_NO = dbo.FILE8_50.DOC_NO INNER JOIN
                      dbo.FILE8_51 ON dbo.FILE8_50.CHARGE = dbo.FILE8_51.CODE
WHERE     (NOT (dbo.FILE8_50.BOX IS NULL))
UNION ALL
SELECT     dbo.FILE8_60H.DOC_NO, dbo.FILE8_60.VALUE AS PLUS, 0 AS MINUS, dbo.FILE8_60.BOX, dbo.FILE8_60H.DATE, 6 AS FLAG, '«Ì—«œ« ' AS DESCA, 
                      dbo.FILE8_61.DESCA AS Expr1
FROM         dbo.FILE8_60H INNER JOIN
                      dbo.FILE8_60 ON dbo.FILE8_60H.DOC_NO = dbo.FILE8_60.DOC_NO LEFT OUTER JOIN
                      dbo.FILE8_61 ON dbo.FILE8_60.CHARGE = dbo.FILE8_61.CODE
WHERE     (NOT (dbo.FILE8_60.BOX IS NULL))
UNION ALL
SELECT     code, 0 AS MINUS, value, no1, date, 7 AS FLAG, ' ÕÊÌ· „‰ Œ“«‰…' AS DESCA, NULL AS Expr1
FROM         dbo.FILE0_51
WHERE     (NOT (no1 IS NULL))
UNION ALL
SELECT     code, value AS PLUS, 0 AS MINSU, no2 AS BOX, date, 8 AS FLAG, ' ÕÊÌ· «·Ì Œ“«‰…' AS DESCA, NULL AS Expr1
FROM         dbo.FILE0_51 AS FILE0_51_1
WHERE     (NOT (no2 IS NULL))
UNION ALL
SELECT     NULL AS Expr1, F_BAL, 0 AS MINUS, CODE, F_DATE, 0 AS FLAG, '—’Ìœ «Ê·' AS DESCA, NULL AS Expr2
FROM         dbo.FILE0_50
UNION ALL
SELECT     dbo.FILE6_20H.DOC_NO, dbo.FILE6_20H.CASH AS PLUS, 0 AS MINUS, dbo.FILE6_20H.BOX, dbo.FILE6_20H.date, 9 AS FLAG, '„»Ì⁄«  ‰ﬁœÌ…' AS DESCA, 
                      FILE3_10_2.DESCA AS Expr1
FROM         dbo.FILE6_20H INNER JOIN
                      dbo.FILE6_20 ON dbo.FILE6_20H.DOC_NO = dbo.FILE6_20.DOC_NO INNER JOIN
                      dbo.FILE3_10 AS FILE3_10_2 ON dbo.FILE6_20H.code = FILE3_10_2.CODE
WHERE     (NOT (dbo.FILE6_20H.BOX IS NULL))
GROUP BY dbo.FILE6_20H.DOC_NO, dbo.FILE6_20H.BOX, dbo.FILE6_20H.CASH, dbo.FILE6_20H.date, FILE3_10_2.DESCA, dbo.FILE6_20H.discount, 
                      dbo.FILE6_20H.tax
UNION ALL
SELECT     dbo.FILE6_10H.DOC_NO, 0 AS PLUS, dbo.FILE6_10H.CASH AS MINUS, dbo.FILE6_10H.BOX, dbo.FILE6_10H.date, 10 AS FLAG, 
                      '„—œÊœ „»Ì⁄«  ‰ﬁœÌ…' AS DESCA, FILE3_10_3.DESCA AS Expr1
FROM         dbo.FILE6_10H INNER JOIN
                      dbo.FILE6_10 ON dbo.FILE6_10H.DOC_NO = dbo.FILE6_10.DOC_NO INNER JOIN
                      dbo.FILE3_10 AS FILE3_10_3 ON dbo.FILE6_10H.code = FILE3_10_3.CODE
WHERE     (NOT (dbo.FILE6_10H.BOX IS NULL))
GROUP BY dbo.FILE6_10H.DOC_NO, dbo.FILE6_10H.BOX, dbo.FILE6_10H.CASH, dbo.FILE6_10H.date, FILE3_10_3.DESCA, dbo.FILE6_10H.discount, 
                      dbo.FILE6_10H.tax
UNION ALL
SELECT     dbo.FILE7_20H.DOC_NO, 0 AS PLUS, dbo.FILE7_20H.CASH AS MINUS, dbo.FILE7_20H.BOX, dbo.FILE7_20H.date, 11 AS FLAG, '„‘ —Ì«  ‰ﬁœÌ…' AS DESCA, 
                      file4_10_2.DESCA AS Expr1
FROM         dbo.FILE7_20H INNER JOIN
                      dbo.FILE7_20 ON dbo.FILE7_20H.DOC_NO = dbo.FILE7_20.DOC_NO INNER JOIN
                      dbo.file4_10 AS file4_10_2 ON dbo.FILE7_20H.code = file4_10_2.CODE
WHERE     (NOT (dbo.FILE7_20H.BOX IS NULL))
UNION ALL
SELECT     dbo.FILE7_30H.DOC_NO, dbo.FILE7_30H.CASH AS PLUS, 0 AS MINUS, dbo.FILE7_30H.BOX, dbo.FILE7_30H.date, 12 AS FLAG, 
                      '„—œÊœ „‘ —Ì«  ‰ﬁœÌ…' AS DESCA, file4_10_3.DESCA AS Expr1
FROM         dbo.FILE7_30H INNER JOIN
                      dbo.FILE7_30 ON dbo.FILE7_30H.DOC_NO = dbo.FILE7_30.DOC_NO INNER JOIN
                      dbo.file4_10 AS file4_10_3 ON dbo.FILE7_30H.code = file4_10_3.CODE
WHERE     (NOT (dbo.FILE7_30H.BOX IS NULL))
GROUP BY dbo.FILE7_30H.DOC_NO, dbo.FILE7_30H.BOX, dbo.FILE7_30H.date, file4_10_3.DESCA, dbo.FILE7_30H.discount, dbo.FILE7_30H.tax, 
                      dbo.FILE7_30H.CASH
UNION ALL
SELECT     dbo.FILE5_20.SER_NO, dbo.FILE5_20.VALUE AS PLUS, 0 AS MINUS, dbo.FILE5_20.BOX, dbo.FILE5_20.DATE_3, 13 AS FLAG, '«Ê—«ﬁ ﬁ»÷' AS DESCA, 
                      FILE3_10_4.DESCA + file4_10_4.DESCA + CASE WHEN (FILE3_10_4.DESCA IS NULL) AND (FILE4_10_4.DESCA IS NULL) 
                      THEN FILE5_20.DESCA ELSE NULL END AS Expr1
FROM         dbo.FILE5_20 LEFT OUTER JOIN
                      dbo.FILE3_10 AS FILE3_10_4 ON FILE3_10_4.CODE = dbo.FILE5_20.code1 LEFT OUTER JOIN
                      dbo.file4_10 AS file4_10_4 ON dbo.FILE5_20.code2 = file4_10_4.CODE
WHERE     (NOT (dbo.FILE5_20.BOX IS NULL)) AND (dbo.FILE5_20.CLOSED = '2')
UNION ALL
SELECT     dbo.file5_30.doc_no, dbo.file5_30.value2 AS PLUS, 0 AS MINUS, dbo.file5_30.box, dbo.file5_30.date, 14 AS FLAG, '„”ÕÊ»« ' AS DESCA, 
                      dbo.file5_00.desca AS Expr1
FROM         dbo.file5_30 INNER JOIN
                      dbo.file5_30H ON dbo.file5_30.doc_no = dbo.file5_30H.doc_no INNER JOIN
                      dbo.file5_00 ON dbo.file5_30.code = dbo.file5_00.code
WHERE     (NOT (dbo.file5_30.box IS NULL)) AND (dbo.file5_30.value2 = 0)
UNION ALL
SELECT     file5_30_1.doc_no, 0 AS PLUS, file5_30_1.value1 AS MINUS, file5_30_1.box, file5_30_1.date, 15 AS FLAG, ' «Ìœ«⁄«  »‰ﬂÌ…' AS DESCA, 
                      file5_00_1.desca AS Expr1
FROM         dbo.file5_30 AS file5_30_1 INNER JOIN
                      dbo.file5_30H AS file5_30H_1 ON file5_30_1.doc_no = file5_30H_1.doc_no INNER JOIN
                      dbo.file5_00 AS file5_00_1 ON file5_30_1.code = file5_00_1.code
WHERE     (NOT (file5_30_1.box IS NULL)) AND (NOT (file5_30_1.value1 = 0))
UNION ALL
SELECT     dbo.FILE5_21.SER_NO, dbo.FILE5_21.VALUE AS PLUS, 0 AS MINUS, dbo.FILE5_21.BOX, dbo.FILE5_21.DATE_3, 16 AS FLAG, '«Ê—«ﬁ œ›⁄' AS DESCA, 
                      file4_10_5.DESCA + CASE WHEN (file4_10_5.DESCA IS NULL) THEN FILE5_21.DESCA ELSE NULL END AS Expr1
FROM         dbo.FILE5_21 LEFT OUTER JOIN
                      dbo.file4_10 AS file4_10_5 ON file4_10_5.CODE = dbo.FILE5_21.code1
WHERE     (NOT (dbo.FILE5_21.BOX IS NULL)) AND (dbo.FILE5_21.CLOSED = '2')
UNION ALL
SELECT     dbo.FILE8_70H.DOC_NO, dbo.FILE8_70.VALUE_P AS PLUS, dbo.FILE8_70.VALUE_M AS MINUS, dbo.FILE8_70.BOX, dbo.FILE8_70H.DATE, 17 AS FLAG, 
                      'Ã«—Ì ‘—ﬂ«¡' AS DESCA, dbo.FILE8_71.DESCA AS CODEDESCA
FROM         dbo.FILE8_70H INNER JOIN
                      dbo.FILE8_70 ON dbo.FILE8_70H.DOC_NO = dbo.FILE8_70.DOC_NO LEFT OUTER JOIN
                      dbo.FILE8_71 ON dbo.FILE8_70.CODE = dbo.FILE8_71.CODE
WHERE     (NOT (dbo.FILE8_70.BOX IS NULL))
