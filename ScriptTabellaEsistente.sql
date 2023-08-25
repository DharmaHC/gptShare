USE [SSM_Data_A1S]
GO

/****** Object:  Table [dbo].[InventoryItems]    Script Date: 25/08/2023 13:23:40 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[InventoryItems](
	[ItemCode] [nvarchar](100) NOT NULL,
	[ItemProviderCode] [varchar](100) NULL,
	[Description] [varchar](100) NULL,
	[CreationDate] [datetime] NOT NULL,
	[IsActive] [bit] NOT NULL,
	[TypeCode] [nvarchar](10) NULL,
	[ProviderId] [int] NULL,
	[StockThresholds] [int] NULL,
	[DefaultOrderQnt] [int] NULL,
	[Price] [money] NULL,
	[Discount] [money] NULL,
	[FinalPrice] [money] NULL,
	[LastOrderDate] [datetime] NULL,
	[ActualStock] [int] NULL,
 CONSTRAINT [PK_InventoryItems] PRIMARY KEY CLUSTERED 
(
	[ItemCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[InventoryItems] ADD  CONSTRAINT [DF_InventoryItems_IsActive]  DEFAULT ((0)) FOR [IsActive]
GO

ALTER TABLE [dbo].[InventoryItems]  WITH CHECK ADD  CONSTRAINT [FK_InventoryItems_InventoryItemTypes] FOREIGN KEY([TypeCode])
REFERENCES [dbo].[InventoryItemTypes] ([TypeCode])
GO

ALTER TABLE [dbo].[InventoryItems] CHECK CONSTRAINT [FK_InventoryItems_InventoryItemTypes]
GO

ALTER TABLE [dbo].[InventoryItems]  WITH CHECK ADD  CONSTRAINT [FK_InventoryItems_InventoryProviders] FOREIGN KEY([ProviderId])
REFERENCES [dbo].[InventoryProviders] ([ProviderId])
GO

ALTER TABLE [dbo].[InventoryItems] CHECK CONSTRAINT [FK_InventoryItems_InventoryProviders]
GO


