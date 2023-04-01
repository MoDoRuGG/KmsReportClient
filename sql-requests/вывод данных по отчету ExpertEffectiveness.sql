SELECT *
  FROM [kms_report].[dbo].[Report_ExpertEffectiveness] a
  JOIN [kms_report].[dbo].[Report_Data] b on a.Id_Report_Data=b.Id
  JOIN [kms_report].[dbo].[Report_Flow] c on b.Id_Flow=c.Id
  WHERE Id_Report_Type = 'Effective'
  Order by a.Id


  SELECT *
  FROM [kms_report].[dbo].[Report_Cadre] a
  JOIN [kms_report].[dbo].[Report_Data] b on a.Id_Report_Data=b.Id
  JOIN [kms_report].[dbo].[Report_Flow] c on b.Id_Flow=c.Id
  WHERE Id_Report_Type = 'Cadre'
  --Order by a.Id


    SELECT *
  FROM [kms_report].[dbo].[Report_Zpz] a
  JOIN [kms_report].[dbo].[Report_Data] b on a.Id_Report_Data=b.Id
  JOIN [kms_report].[dbo].[Report_Flow] c on b.Id_Flow=c.Id
  WHERE Id_Report_Type = 'Zpz'