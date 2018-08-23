//---------------------------------------------------------------------------

#ifndef XlsToCSVThreadH
#define XlsToCSVThreadH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include "uMain.h"
//---------------------------------------------------------------------------
class TXlsToCSVThread : public TThread
{
private:
  String m_sFile;
protected:
	void __fastcall Execute();
public:
	__fastcall TXlsToCSVThread(bool CreateSuspended, String sFile);daksdakshd
};
//---------------------------------------------------------------------------
#endif
