USE [entnavistar]
GO

/****** Object:  Table [dbo].[MC_InterfaceLog]    Script Date: 9/22/2016 10:25:26 AM ******/
DROP TABLE [dbo].[MC_InterfaceLog]
GO

/****** Object:  Table [dbo].[MC_InterfaceLog]    Script Date: 9/22/2016 10:25:26 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MC_InterfaceLog](
	[PK] [int] IDENTITY(1,1) NOT NULL,
	[ProcessDate] [datetime] NULL,
	[ImportID] [varchar](38) NULL,
	[VehicleID] [varchar](250) NULL,
	[RecordNumber] [int] NULL,
	[MCRecordPK] [int] NULL,
	[ErrorMessage] [varchar](7000) NULL,
	[Processed] [char](1) NULL,
	[RecordData] [varchar](7000) NULL,
	[Hours] [varchar](150) NULL,
	[Miles] [varchar](25) NULL,
 CONSTRAINT [PK_MC_InterfaceLog] PRIMARY KEY CLUSTERED 
(
	[PK] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

