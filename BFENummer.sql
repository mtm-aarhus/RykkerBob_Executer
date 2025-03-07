SELECT TOP (100)

	   [EjendePersonPersonNR]
      ,[BFEnummer]
      ,[EjendeVirksomhedCVRnr]
      ,[EjerType]
      ,[Navn]
      ,[Kommunekode]
	  ,EjerStatusKode_T
	  ,PrimaerKontakt_T
	  ,NavnJusteret
	  ,EjersAdresseDanmark
	  ,Beskyttelse
      ,Beskyttelse_T

  FROM [EJF].[EjendomsejerView]

    Where BFEnummer = 'Ejendomsnummer'
