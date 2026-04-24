USE [OLYMPIC]
GO

/****** Object:  StoredProcedure [dbo].[sp_paid_card]    Script Date: 29/01/2024 02:19:49 Õ ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[sp_paid_card]
(
	@SEASON INT,
	@DATE1 DATETIME = NULL,
	@DATE2 DATETIME = NULL,
	@CODE INT = NULL,
	@CODE1 INT = NULL,
	@CODE2 INT = NULL,
	@TYPE INT = NULL,
	@PATH NVARCHAR(100) = NULL
)
AS
BEGIN
	SELECT vw_members_all.[MEMBER], 
	vw_members_all.CODE, 
	vw_members_all.DESCA, 
	RELATION_CODES.DESCA as relation_desca, 
	vw_last_card.DOC_NO,
    convert(varchar(10),vw_last_card.[Date],111) as [DATE],
	FILE6_20H.YEARS_DESCA,
	vw_members_all.DESCA_MEMBER, 
	vw_members_all.DATE_BIRTH,
	vw_members_all.TITLE,
	vw_members_all.GENDER,
	vw_members_all.RELATION,
	vw_members_all.HANDI,
	vw_members_all.ID	
	FROM  vw_members_all INNER JOIN vw_last_card ON vw_members_all.MEMBER = vw_last_card.CODE 
	inner JOIN FILE1_10 ON vw_members_all.member = file1_10.code 
	inner join file6_20h on file6_20h.doc_no = vw_last_card.DOC_NO 
	left join relation_codes on vw_members_all.RELATION = relation_codes.code
	WHERE vw_last_card.YEAR_CODE <= @SEASON
	AND ((not vw_members_all.code is null) or file1_10.died = 0) 
	AND (@CODE IS NULL OR vw_members_all.[MEMBER] = @CODE)
	AND (@CODE1 IS NULL OR vw_members_all.[MEMBER] >= @CODE1)
	AND (@CODE2 IS NULL OR vw_members_all.[MEMBER] <= @CODE2)
	AND (@DATE1 IS NULL OR FILE6_20H.[DATE] >= @DATE1)
	AND (@DATE2 IS NULL OR FILE6_20H.[DATE] <= @DATE2)
	AND (@TYPE IS NULL OR FILE1_10.[TYPE] = @TYPE)
	AND (@PATH IS NULL OR 
	[dbo].[fn_file_exist](@PATH + CAST(vw_members_all.[MEMBER] AS VARCHAR(10)) + 
	CASE WHEN vw_members_all.[CODE] IS NULL then '' ELSE '-' END + COALESCE(CAST(vw_members_all.[CODE] AS VARCHAR(10)),'') + '.jpg') = 1)
	ORDER BY vw_members_all.MEMBER, vw_members_all.CODE
END





GO


