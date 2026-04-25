---
title: Cannot open database <name>. It may not be a database that your application recognizes, or the file may be corrupt. (Error 3049)
keywords: jeterr40.chm5003049
f1_keywords:
- jeterr40.chm5003049
ms.assetid: 5441640a-c2e9-ac40-f7d7-1b1a216c9fd8
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Cannot open database \<name\>. It may not be a database that your application recognizes, or the file may be corrupt. (Error 3049)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- You tried to open a non-Microsoft Access database engine database, such as a Btrieve, dBASE, or Paradox database or table. Instead, link the database or table to an existing Microsoft Jet database.
    
- You tried to import or link a Paradox database or table, and the associated .px file could not be found. Make sure the .px file is the same directory as the Paradox database or table, and then try the operation again.
    
- If the specified database is a Microsoft Jet database, it is corrupted. You should attempt to repair the database. If the database cannot be repaired, restore the database from a backup copy, or create a new database.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
