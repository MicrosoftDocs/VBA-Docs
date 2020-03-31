---
title: Input and InputB functions (differences in string function operations)
ROBOTS: INDEX
ms.prod: access
ms.assetid: 49389ce7-dc9b-3014-ec09-5c444d60b8bf
ms.date: 06/08/2019
localization_priority: Normal
---


# Input and InputB functions (differences in string function operations)

**Applies to:** Access 2013 | Access 2016

The memory storage formats for text differ between Visual Basic for Applications (VBA) code and Access Basic code. (Access Basic was used in early versions of Microsoft Access.) Text is stored in ANSI format within Access Basic code and in Unicode format in Visual Basic. This topic discusses one potential issue when handling strings in the current version of Microsoft Access. 

For more information, see [Differences in String Function Operations](https://msdn.microsoft.com/library/40ce2b9a-cac6-589e-2b5e-d63be37efeee%28Office.15%29.aspx).

The **Input** function in Microsoft Access converts the number of characters designated when the text is read from the file into a Unicode string and reads them as variables. The **InputB** function, on the other hand, assumes the data to be binary data and stores it as variables without converting it. If the **InputB** function is used when reading a file stored in a fixed length field, the fixed byte length data must be converted once it is read.

```vb
Open "Data.Dat" For Input As 1 
dat1 = StrConv(InputB(10, 1), vbUnicode) 
dat2 = StrConv(InputB(10, 1), vbUnicode) 
dat3 = StrConv(InputB(10, 1), vbUnicode) 
 
===DATA.DAT 
123456789012345678901234567 
Name Address Telephone
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]