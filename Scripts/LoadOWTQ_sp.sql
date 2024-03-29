USE [FT_AppMidware]
GO
/****** Object:  StoredProcedure [dbo].[LoadOWTR_sp]    Script Date: 11/13/2020 11:36:07 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

create PROCEDURE [dbo].[LoadOWTQ_sp]
	
AS

select T0.guid as [Key],T2.transdate as [DocDate], T1.fromwarehouse as [WarehouseFrom],T1.towarehouse as [WarehouseTo],
T1.ItemCode, T1.quantity as [Quantity]

  
from zmwRequest T0
inner join zmwInventoryRequest T1 on T0.guid = T1.Guid
left join zmwInventoryRequestHead T2 on T0.guid = T2.Guid
where status ='ONHOLD' and T0.request = 'Create Inventory Request'
--and T1.itemcode = 'B10000'