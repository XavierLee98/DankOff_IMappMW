USE [FT_AppMidware]
GO
/****** Object:  StoredProcedure [dbo].[LoadOIGN_sp]    Script Date: 11/2/2020 9:57:55 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

create PROCEDURE [dbo].[LoadOIGN_sp]
	
AS

update T0 set lastErrorMessage ='Warehouse is missing.'
from zmwRequest T0
inner join zmwGRPO T1 on T0.guid = T1.Guid
where status ='ONHOLD' and T0.request = 'Create GR' and isnull(T1.Warehouse,'') = ''


select T0.guid as [Key],T0.requestTime as [DocDate], T1.SourceCardCode as [CardCode],
T1.SourceDocBaseType as [BaseType], T1.SourceBaseEntry as [BaseEntry], T1.SourceBaseLine as [BaseLine],
T1.ItemCode, T1.qty as [Quantity], isnull(T1.Warehouse,'') as [whscode],
T2.DocSeries, T2.Ref2,T2.Comments,T2.JrnlMemo,T2.NumAtCard
  
from zmwRequest T0
inner join zmwGRPO T1 on T0.guid = T1.Guid
left join zmwDocHeaderField T2 on T0.guid = T2.Guid
where status ='ONHOLD' and T0.request = 'Create GR' and isnull(T1.Warehouse,'') <> ''
