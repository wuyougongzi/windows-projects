#ifndef DRAGDROPHELPER_H
#define DRAGDROPHELPER_H

#include "dataobject.h"
#include "dropsource.h"

class CDragDropHelper
{
private:
	CDragDropHelper();
	~CDragDropHelper();

public:
	static CDragDropHelper* GetInstance();

public:
	void DoDragDrop(HGLOBAL hGlobal);
	void DoDragDrop(HGLOBAL hGlobal, HBITMAP hbitmap);
	
private:
	static CDragDropHelper*	m_pInstance;
};

#endif 