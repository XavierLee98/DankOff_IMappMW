USE [FT_AppMidware]
GO

/****** Object:  Table [dbo].[ft_Errors]    Script Date: 10/31/2020 6:12:33 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ft_Errors](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Module] [nvarchar](50) NULL,
	[ErrMsg] [nvarchar](1000) NULL,
	[CreateDate] [datetime] NULL,
 CONSTRAINT [PK_ft_Errors] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


