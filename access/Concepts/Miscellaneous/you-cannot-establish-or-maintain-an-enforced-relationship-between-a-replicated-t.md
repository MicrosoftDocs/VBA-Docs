---
title: You cannot establish or maintain an enforced relationship between a replicated table and a local table. (Error 3453)
keywords: jeterr40.chm5003453
f1_keywords:
- jeterr40.chm5003453
ms.prod: access
ms.assetid: 1bd3124e-452f-4cd7-8c71-dbc3267e63a8
ms.date: 06/08/2017
localization_priority: Normal
---


# You cannot establish or maintain an enforced relationship between a replicated table and a local table. (Error 3453)

  

**Applies to:** Access 2013 | Access 2016

You are attempting to establish or maintain an enforced relationship between a replicated table and a non-replicated table. Replication does not allow an enforced relationship between:



- A replicated table and a local table.
    
- Two local tables that you are making replicable.
    
- Two tables with different  **KeepLocal** property settings.
    

Delete the relationship between the two tables before proceeding.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]