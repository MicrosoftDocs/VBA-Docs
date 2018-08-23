---
title: In operator without (). (Error 2429)
keywords: jeterr40.chm5002429
f1_keywords:
- jeterr40.chm5002429
ms.prod: access
ms.assetid: 40f2356c-f891-1d90-17be-ace51c989357
ms.date: 06/08/2017
---


# In operator without (). (Error 2429)

  

**Applies to:** Access 2013 | Access 2016

When coding an SQL statement that includes the  **In** operator, you must surround the list of items to test with parentheses. For example, to see if a value is one of a set of values, you could use the following code in the WHERE clause of an SQL query:




```vb
WHERE Region In ('TX', 'CA', 'WA')

```

This code tests to see if the Region field contains any of the above abbreviations, which represent Texas, California, and Washington.

## See also

- [Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/en-us/msoffice/forum?page=1&;tab=question&;status=all&;auth=1)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)