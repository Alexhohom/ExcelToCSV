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
  #define PG OlePropertyGet    //��ȡ��������
  #define PS OlePropertySet    //���ö�������
  #define FN OleFunction       //���ö��󷽷�
  #define PR OleProcedure
  Variant 	ExcelApp;      //Excel�е�Application����,ExcelӦ�ó������
  Variant 	WorkBooks;     //Excel�е�Workbooks����,Excel����������
  Variant 	WorkBook1;
  Variant 	Sheets;        //Excel�е�Sheets����Excel�еĹ��������
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
      MainForm->AddLog(L"Ŀ¼�����쳣");
      return;
    }
  }
  __int64 	iTimeOpen = GetTickCount();

  try
  {
    ExcelApp = Variant::CreateObject("Excel.Application");

    if (ExcelApp.IsNull() || ExcelApp.IsEmpty())
    {
      MainForm->AddLog(L"ExcelӦ�ÿ�ʧ��");
      return;
    }
  }
  catch (...)
  {

    MainForm->AddLog(L"ExcelӦ�û�ȡʧ��");
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
    sLog.sprintf(L" openExcelSpend  %lld��", (GetTickCount()-iTimeOpen)/1000);
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
      sLog.sprintf(L"����: %d �����: %d", iSheetCount, i);
      MainForm->AddLog(sLog);
    }
    String sSaveLog;
    sSaveLog.sprintf(L" SaveExcelSpend  %lld�� ������%d�� Ŀ¼:%s",
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
    MainForm->AddLog(L"�쳣");
    return;
  }
#endif
}
//---------------------------------------------------------------------------
