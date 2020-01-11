---
title: Size property (ADO stream)
ROBOTS: INDEX
ms.prod: access
ms.assetid: deb84313-36d1-fa49-e4cd-daecab96f343
ms.date: 06/08/2017
localization_priority: Normal
---


# Size property (ADO stream)

**Applies to:** Access 2013 | Access 2016

Indicates the size of the stream in number of bytes.

## Return values

Returns a **Long** value that specifies the size of the stream in number of bytes. The default value is the size of the stream, or -1 if the size of the stream is not known.


## Remarks

**Size** can be used only with open [Stream](https://msdn.microsoft.com/library/d49b1514-e0b4-0aca-d5c2-8266f3f4fe65%28Office.15%29.aspx) objects.

> [!NOTE] 
> Any number of bits can be stored in a **Stream** object, limited only by system resources. If the **Stream** contains more bits than can be represented by a **Long** value, **Size** is truncated and therefore does not accurately represent the length of the **Stream**.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]