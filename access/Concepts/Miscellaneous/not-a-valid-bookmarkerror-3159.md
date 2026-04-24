---
title: Not a valid bookmark. (Error 3159)
keywords: jeterr40.chm5003159
f1_keywords:
- jeterr40.chm5003159
ms.assetid: 99e8083c-d098-916f-3160-d9787e354216
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Not a valid bookmark. (Error 3159)

  

**Applies to:** Access 2013 | Access 2016

You tried to set a bookmark to an invalid string.

This error can occur if you set the **Bookmark** property to a string that is invalid or was not saved from previously reading a **Bookmark** property. For example, the following code produces this error:



```vb
Sub SetBookmark() 
    Dim dbs As Database 
    Dim rstEmployees As Recordset 
    Dim strPlaceholder As String 

    Set dbs = OpenDatabase("Northwind.mdb") 

    Set rstEmployees = _ 
        dbs.OpenRecordset _
        ("Employees", dbOpenDynaset) 

    strPlaceholder = "1" 

    rstEmployees.Bookmark = strPlaceholder    ' Not a valid bookmark. 
End Sub
```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
