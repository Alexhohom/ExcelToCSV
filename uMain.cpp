//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TMainForm *MainForm;
//---------------------------------------------------------------------------
__fastcall TMainForm::TMainForm(TComponent* Owner)
	: TForm(Owner)
{
  Timer1->Interval = 500;
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::btnToCSVClick(TObject *Sender)
{
  String    sFilePath;

  if (OpenDialog1->Execute())
  {
    sFilePath = OpenDialog1->FileName;
  }
  else
  {
    AddLog(L"文件获取失败");
    return;
  }

 
  TXlsToCSVThread *tThread = new TXlsToCSVThread(false, sFilePath);

}
//---------------------------------------------------------------------------
void __fastcall TMainForm::AddLog(String sLog)
{
  String sOutLog;
  sOutLog.sprintf(L"%s  %s", FormatDateTime("yyyy-mm-dd hh:mm:ss", Now()), sLog);

  list1->Items->Add(sOutLog);
  int iListPos = list1->Items->Count;
  list1->ItemIndex = iListPos -1;
}
void __fastcall TMainForm::Timer1Timer(TObject *Sender)
{
  String sLog;
  if (iState == 1)
  {

  	sLog.sprintf(L" 正在打开文件,请稍等...", iState);
    AddLog(sLog);
  }


}
//---------------------------------------------------------------------------

