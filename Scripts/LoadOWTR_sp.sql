USE [FT_AppMidware]
GO
/****** Object:  StoredProcedure [dbo].[LoadOPDN_sp]    Script Date: 11/12/2020 3:39:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

create PROCEDURE [dbo].[LoadOWTR_sp]
	
AS

select T0.guid as [Key],T2.docdate as [DocDate], T2.taxdate as [TaxDate],T1.fromwhscode as [WarehouseFrom],T1.towhscode as [WarehouseTo],
T2.Comments,T2.JrnlMemo,
T1.SourceDocBaseType as [BaseType], T1.SourceBaseEntry as [BaseEntry], T1.SourceBaseLine as [BaseLine],
T1.ItemCode, T1.actualreceiptqty as [Quantity]

  
from zmwRequest T0
inner join zmwTransferDocDetails T1 on T0.guid = T1.Guid
left join zmwTransferDocHeader T2 on T0.guid = T2.Guid
where status ='ONHOLD' and T0.request = 'Create Transfer1'
