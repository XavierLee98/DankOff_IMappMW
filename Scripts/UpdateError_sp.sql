USE [FT_AppMidware]
GO

/****** Object:  StoredProcedure [dbo].[UpdateError_sp]    Script Date: 10/31/2020 6:13:15 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



create PROCEDURE [dbo].[UpdateError_sp]
(
	@MODULE	NVARCHAR(100),
	@ERRMSG	NVARCHAR(100)
)
AS
	INSERT INTO FT_ERRORS
	(
		MODULE,
		ERRMSG,
		CREATEDATE
	)
	VALUES
	(
		@MODULE,
		@ERRMSG,
		GETDATE()
	)

GO


