USE [FT_AppMidware]
GO
/****** Object:  StoredProcedure [dbo].[LoadBinDetails_sp]    Script Date: 11/11/2020 3:41:13 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


create PROCEDURE [dbo].[LoadBinDetails_sp]
	(
		@request	nvarchar(100)=null
	)
AS

select T2.*

from zmwRequest T0
inner join zmwGRPO T1 on T0.guid = T1.Guid
left join zmwItemBin T2 on T0.guid = T2.guid and T1.ItemCode = T2.ItemCode
where status ='ONHOLD' and T0.request = @request and T2.guid is not null and isnull(T2.BinCode,'') <> ''
--and T2.ItemCode = 'B10000'
