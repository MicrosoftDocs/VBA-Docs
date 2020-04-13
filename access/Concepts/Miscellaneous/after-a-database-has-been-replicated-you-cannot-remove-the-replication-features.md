---
title: After a database has been replicated, you cannot remove the replication features from the database. (Error 3459)
keywords: jeterr40.chm5003459
f1_keywords:
- jeterr40.chm5003459
ms.prod: access
ms.assetid: 06429f0d-7b78-d7e7-3c67-23142282b0ca
ms.date: 06/08/2019
localization_priority: Normal
---


# After a database has been replicated, you cannot remove the replication features from the database. (Error 3459)

  

**Applies to:** Access 2013 | Access 2016

You cannot remove the replication features from a database that has been replicated, either by setting its **Replicable** property to "T" or by converting it with the Microsoft Windows Briefcase, Microsoft Access, or the Replication Manager. Using DAO to set the database's **Replicable** property to "F" or any other value has no effect on the replicated database. If you need to use a copy of the database that does not have the properties, fields, tables, and other characteristics associated with replication, open the backup copy of the database made before the database was first converted.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]