USE [FT_AppMidware]
GO
/****** Object:  StoredProcedure [dbo].[LoadDetails_sp]    Script Date: 11/11/2020 3:41:33 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


create PROCEDURE [dbo].[LoadDetails_sp]
	(
	@request	nvarchar(100)=null
	)
AS

select T2.guid,T2.itemcode, sum(T2.quantity) as [quantity],T2.SerialNumber, T2.BatchNumber, T2.BatchAttr1,T2.BatchAttr2,convert(varchar(10),T2.BatchAdmissionDate,126) as [BatchAdmissionDate]

from zmwRequest T0
inner join zmwGRPO T1 on T0.guid = T1.Guid
left join zmwItemBin T2 on T0.guid = T2.guid and T1.ItemCode = T2.ItemCode
where status ='ONHOLD' and T0.request =@request and T2.guid is not null --and isnull(T2.BatchNumber,'') <> ''
--and T2.ItemCode = 'B10000'

group by T2.guid,T2.itemcode, T2.SerialNumber, T2.BatchNumber, T2.BatchAttr1,T2.BatchAttr2,convert(varchar(10),T2.BatchAdmissionDate,126)
