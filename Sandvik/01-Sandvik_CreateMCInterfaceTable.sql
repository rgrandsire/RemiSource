USE [entSandvik]
GO

/****** Object:  Table [dbo].[MC_InterfaceLog]    Script Date: 11/17/2016 3:08:23 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MC_InterfaceLog](
	[PK] [int] IDENTITY(1,1) NOT NULL,
	[ProcessDate] [datetime] NULL,
	[FileName] [varchar](250) NULL,
	[RecordNumber] [int] NULL,
	[ErrorMessage] [varchar](7000) NULL,
	[Processed] [char](1) NULL,
 CONSTRAINT [PK_MC_InterfaceLog] PRIMARY KEY CLUSTERED 
(
	[PK] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

