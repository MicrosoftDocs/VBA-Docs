---
title: The width of a Unicode text column must be an even number of bytes. (Error 3702)
keywords: jeterr40.chm5003702
f1_keywords:
- jeterr40.chm5003702
ms.prod: access
ms.assetid: 029d1aa7-3a75-8aa6-3255-c39c5a62d358
ms.date: 06/08/2017
localization_priority: Normal
---


# The width of a Unicode text column must be an even number of bytes. (Error 3702)

  

**Applies to:** Access 2013 | Access 2016

The database stores data using the Unicode character set, in which each character is a word (two-byte) value. Columns for text must be declared to be an even number of bytes wide so that the columns can hold an exact number of Unicode characters.



- If this error message appears during database creation, the program that is attempting to create an incorrect database should be corrected.
    
- If this error message appears during normal run time, it is probable that the database has become corrupted and needs to be compacted or repaired.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]