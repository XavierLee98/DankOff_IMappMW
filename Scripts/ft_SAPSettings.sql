USE [FT_AppMidware]
GO

/****** Object:  Table [dbo].[ft_SAPSettings]    Script Date: 10/31/2020 12:28:00 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ft_SAPSettings](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[UserName] [nvarchar](20) NULL,
	[DBUser] [nvarchar](20) NULL,
	[DBPass] [nvarchar](20) NULL,
	[SAPCompany] [nvarchar](20) NULL,
	[DBType] [int] NULL,
	[LicenseServer] [nvarchar](50) NULL,
	[Server] [nvarchar](50) NULL,
	[SAPUser] [nvarchar](20) NULL,
	[SAPPass] [nvarchar](20) NULL,
 CONSTRAINT [PK_ft_SAPSettings] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


