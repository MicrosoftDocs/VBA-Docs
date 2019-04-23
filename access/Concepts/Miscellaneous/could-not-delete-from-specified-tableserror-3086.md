---
title: Could not delete from specified tables. (Error 3086)
keywords: jeterr40.chm5003086
f1_keywords:
- jeterr40.chm5003086
ms.prod: access
ms.assetid: c0b7ba20-7b1a-a6de-b2e0-8ec095a0e448
ms.date: 06/08/2017
localization_priority: Normal
---


# Could not delete from specified tables. (Error 3086)

  

**Applies to:** Access 2013 | Access 2016

You tried to delete data from one or more tables, but the deletion could not be completed.

Possible causes:


- You do not have permission to modify the table. To change your permissions assignments, see your system administrator or the table's creator.
    
- The database was opened for read-only access. The database is read-only for one of these reasons:
    
    
    
      - You used the  **OpenDatabase** method and opened the database for read-only access.
    
  - The database file is defined as read-only in the database server operating system or by your network.
    
  - In a network environment, you do not have write privileges for the database file.
    
  - In Microsoft Visual Basic, you used the  **Data** control and set the **ReadOnly** property to **True**.
    

    
    

To delete the data, close the database, resolve the read-only condition, and then reopen the file for read/write access.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
