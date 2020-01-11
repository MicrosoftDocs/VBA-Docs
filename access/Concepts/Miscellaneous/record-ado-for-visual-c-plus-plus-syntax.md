---
title: Record (ADO for Visual C++ syntax)
ms.prod: access
ms.assetid: e9a1300e-e2d8-7ad9-e0d6-61be720b83af
ms.date: 10/12/2018
localization_priority: Normal
---


# Record (ADO for Visual C++ syntax)

**Applies to:** Access 2013 | Access 2016

## Methods

[Cancel](https://msdn.microsoft.com/library/747edc04-a5cc-3631-2d0b-82e7e41a76b7%28Office.15%29.aspx)(void) 

[Close](https://msdn.microsoft.com/library/26a7cced-ebeb-70be-f5de-96a35711bc37%28Office.15%29.aspx)(void) 

[CopyRecord](https://msdn.microsoft.com/library/724e4358-f216-8e47-5bab-c72770ece5a4%28Office.15%29.aspx)(BSTR _Source_, BSTR _Destination_, BSTR _UserName_, BSTR _Password_, CopyRecordOptionsEnum _Options_, VARIANT_BOOL _Async_, BSTR _*pbstrNewURL_ ) 

[DeleteRecord](https://msdn.microsoft.com/library/ba71187f-e580-bba8-f41b-bedfa0bc2b04%28Office.15%29.aspx)(BSTR _Source_, VARIANT_BOOL _Async_ ) 

[GetChildren](https://msdn.microsoft.com/library/998cf640-ffc7-51e1-4d1e-4797f7cdea4a%28Office.15%29.aspx)(_ADORecordset * _*ppRSet_ ) 

[MoveRecord](https://msdn.microsoft.com/library/efc341a2-0e08-a838-5925-8d4c46377e48%28Office.15%29.aspx)(BSTR _Source_, BSTR _Destination_, BSTR _UserName_, BSTR _Password_, MoveRecordOptionsEnum _Options_, VARIANT_BOOL _Async_, BSTR _*pbstrNewURL_ ) 

[Open](https://msdn.microsoft.com/library/ba71c5c7-326e-d3b6-0e74-e8343ee6896f%28Office.15%29.aspx)(VARIANT _Source_, VARIANT _ActiveConnection_, ConnectModeEnum _Mode_, RecordCreateOptionsEnum _CreateOptions_, RecordOpenOptionsEnum _Options_ BSTR _UserName_, BSTR _Password_ )

## Properties

[get_ActiveConnection](https://msdn.microsoft.com/library/5501b2d7-b62c-5fff-1edd-2b7efb3f8c4a%28Office.15%29.aspx)(VARIANT  _*pvar_ ) **put_ActiveConnection** (BSTR _bstrConn_ ) **putref_ActiveConnection** (ADOConnection _*Con_ ) 

[get_Fields](https://msdn.microsoft.com/library/029aa738-8726-54a6-1813-b152813948bc%28Office.15%29.aspx)(ADOFields * _*ppFlds_ ) 

[get_Mode](https://msdn.microsoft.com/library/62086f4f-8624-16c4-dae1-a17475d1864d%28Office.15%29.aspx)(ConnectModeEnum  _*pMode_ ) **put_Mode** (ConnectModeEnum _Mode_ ) 

[get_ParentURL](https://msdn.microsoft.com/library/ec7ec476-6f9e-8486-fe02-74995975df5c%28Office.15%29.aspx)(BSTR  _*pbstrParentURL_ ) 

[get_RecordType](https://msdn.microsoft.com/library/a42001a6-7312-162d-dd71-c82f8c9d527f%28Office.15%29.aspx)(RecordTypeEnum  _*pType_ ) 

[get_Source](https://msdn.microsoft.com/library/f36f0f5f-4493-d8c5-db4b-c72f5031bcb3%28Office.15%29.aspx)(VARIANT  _*pvar_ ) **put_Source** (BSTR _Source_ ) **putref_Source** (IDispatch _*Source_ ) 

[get_State](https://msdn.microsoft.com/library/ade0a50c-e2d8-23ac-4ea9-b012fedcd5db%28Office.15%29.aspx)(ObjectStateEnum  _*pState_ )

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]