USE [FT_AppMidware]
GO
/****** Object:  StoredProcedure [dbo].[LoadDetails_sp]    Script Date: 11/12/2020 3:48:18 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


Create PROCEDURE [dbo].[LoadTransferDetails_sp]
	(
	@request	nvarchar(100)=null
	)
AS

select T2.guid,T2.itemcode, sum(T2.qty) as [quantity],T2.Serial as [serialnumber], T2.Batch as [batchnum], 
T2.internalserialnumber,T2.manufacturerserialnumber, T2.binabs

from zmwRequest T0
inner join zmwTransferDocDetails T1 on T0.guid = T1.Guid
left join zmwTransferDocDetailsBin T2 on T0.guid = T2.guid and T1.ItemCode = T2.ItemCode
where status ='ONHOLD' and T0.request =@request and T2.guid is not null --and isnull(T2.BatchNumber,'') <> ''
--and T2.ItemCode = 'B10000'

group by T2.guid,T2.itemcode, T2.Serial, T2.Batch, T2.internalserialnumber,T2.manufacturerserialnumber, T2.binabs
