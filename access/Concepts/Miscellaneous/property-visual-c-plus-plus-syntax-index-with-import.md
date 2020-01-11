---
title: Property (Visual C++ syntax index with import)
ms.prod: access
ms.assetid: 3649db1b-ab2f-2767-a8b3-a146720217c0
ms.date: 10/12/2018
localization_priority: Normal
---


# Property (Visual C++ syntax index with #import)

**Applies to:** Access 2013 | Access 2016

## Properties

```cs
long GetAttributes( ); 
void PutAttributes( long plAttributes ); 
__declspec(property(get=GetAttributes,put=PutAttributes)) long 
 Invalid DDUE based on source, error:link not allowed in code, link filename:adproattributes_HV10294099.xml; 
 
_bstr_t GetName( ); 
__declspec(property(get=GetName)) _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdproname_HV10294535.xml; 
 
enum DataTypeEnum GetType( ); 
__declspec(property(get=GetType)) enum DataTypeEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprotype_HV10294866.xml; 
 
_variant_t GetValue( ); 
void PutValue( const _variant_t & pval ); 
__declspec(property(get=GetValue,put=PutValue)) _variant_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdprovalue_HV10294920.xml; 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]