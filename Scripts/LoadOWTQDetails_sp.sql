USE [FT_AppMidware]
GO
/****** Object:  StoredProcedure [dbo].[LoadOPDNDetails_sp]    Script Date: 10/31/2020 5:51:27 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

create PROCEDURE [dbo].[LoadOWTQDetails_sp]
	
AS

select T2.*

from zmwRequest T0
inner join zmwGRPO T1 on T0.guid = T1.Guid
left join zmwItemBin T2 on T0.guid = T2.guid and T1.ItemCode = T2.ItemCode
where status ='ONHOLD' and T0.request = 'Create Inventory Request'
