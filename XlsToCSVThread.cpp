//---------------------------------------------------------------------------

#include <System.hpp>
#pragma hdrstop

#include "XlsToCSVThread.h"
#pragma package(smart_init)
//---------------------------------------------------------------------------

//   Important: Methods and properties of objects in VCL can only be
//   used in a method called using Synchronize, for example:
//
//      Synchronize(&UpdateCaption);
//
//   where UpdateCaption could look like:
//
//      void __fastcall TXlsToCSVThread::UpdateCaption()
//      {
//        Form1->Caption = "Updated in a thread";
//      }
//---------------------------------------------------------------------------

__fastcall TXlsToCSVThread::TXlsToCSVThread(bool CreateSuspended, String sFile)
	: TThread(CreateSuspended)
{
  m_sFile = sFile;
}
//---------------------------------------------------------------------------
void __fastcall TXlsToCSVThread::Execute()
{
	//---- Place thread code here ----
  //MainForm->AddLog("test");
#if 1
  #define PG OlePropertyGet    //获取对象属性
  #define PS OlePropertySet    //设置对象属性
  #define FN OleFunction       //调用对象方法
  #define PR OleProcedure
  Variant 	ExcelApp;      //Excel中的Application对象,Excel应用程序对象
  Variant 	WorkBooks;     //Excel中的Workbooks对象,Excel工作簿对象
  Variant 	WorkBook1;
  Variant 	Sheets;        //Excel中的Sheets对象，Excel中的工作表对象
  Variant 	Sheet1;
  String sPath, sFile, sPathFileOut, sFileFormat;
  sPath = ExtractFilePath(m_sFile);
  sFile = ExtractFileName(m_sFile);
  sFileFormat = ExtractFileExt(m_sFile);
  int pos = sFile.Pos(sFileFormat);
  sFile = sFile.SubString(1, pos-1);
  sPathFileOut = sPath + sFile +"\\";
  if (!DirectoryExists(sPathFileOut))
  {
    if (!ForceDirectories(sPathFileOut))
    {
      MainForm->AddLog(L"目录创建异常");
      return;
    }
  }
  __int64 	iTimeOpen = GetTickCount();

  try
  {
    ExcelApp = Variant::CreateObject("Excel.Application");

    if (ExcelApp.IsNull() || ExcelApp.IsEmpty())
    {
      MainForm->AddLog(L"Excel应用开失败");
      return;
    }
  }
  catch (...)
  {

    MainForm->AddLog(L"Excel应用获取失败");
    return;
  }
  try
  {
    ExcelApp.PS("Visible", false);
    MainForm->iState = 1;
    ExcelApp.PG("WorkBooks").PR("Open", WideString(m_sFile));
    MainForm->iState = 0;
    WorkBooks = ExcelApp.PG("ActiveWorkbook");

    #if 1
    String sLog;
    sLog.sprintf(L" openExcelSpend  %lld秒", (GetTickCount()-iTimeOpen)/1000);
    MainForm->AddLog(sLog);
    #endif
    int iSheetCount = WorkBooks.PG("Sheets").PG("Count");
    sFile += (FormatDateTime("yymmddhhmmss", Now())+ "_");
    for (int i = 1; i <= iSheetCount; i++)
    {
      Sheet1 = WorkBooks.PG("Sheets", i);
      String sPathFile;

      sPathFile = sPathFileOut + sFile + IntToStr(i) + ".csv";
      Sheet1.OleProcedure("SaveAs", WideString(sPathFile),6);
      sLog.sprintf(L"总数: %d 已完成: %d", iSheetCount, i);
      MainForm->AddLog(sLog);
    }
    String sSaveLog;
    sSaveLog.sprintf(L" SaveExcelSpend  %lld秒 工作表%d个 目录:%s",
    					((GetTickCount()-iTimeOpen)/1000),
                        iSheetCount, sPathFileOut);
    MainForm->AddLog(sSaveLog);
    theEnd:
    ExcelApp.PG("ActiveWorkbook").PS("Saved", true);
    ExcelApp.FN("Quit");
    return ;
  }
  catch(...)
  {
    ExcelApp.PG("ActiveWorkbook").PS("Saved", true);
    ExcelApp.FN("Quit");
    MainForm->AddLog(L"异常");
    return;
  }
#endif
}
//---------------------------------------------------------------------------
