#ifndef SERVER_MATH_H__
#define SERVER_MATH_H__

#include "IMath.h"


//声明数组引用计数全局变量
extern long g_lObjs;
extern long g_lLocks;
//定义宿主引用计数全局变量
// long g_lObjs = 0;
// long g_lLocks = 0;


class Math : public IMath, IAdvancedMath
{
public:
	Math();
	~Math();

public:
	//IunKnown
	STDMETHOD(QueryInterface(REFIID, void**));
	STDMETHOD_(ULONG, AddRef());
	STDMETHOD_(ULONG, Release());

	//IMath
	STDMETHOD(Add(long, long, long&));
	STDMETHOD(Subtract(long, long, long&));

	//IAdvancedMath
	STDMETHOD(Factorial(short, long&));
	STDMETHOD(Fibonacci(short, long&));

private:
	long	m_lRef;
};



class MathClassFartory : public IClassFactory
{
public:
	MathClassFartory();
	~MathClassFartory();

public:
	//IUnKnown
	STDMETHOD(QueryInterface(REFIID, void**));
	STDMETHOD_(ULONG, AddRef());
	STDMETHOD_(ULONG, Release());

	//IClassFractory
	STDMETHOD(CreateInstance)(LPUNKNOWN, REFIID, void**);
	STDMETHOD(LockServer)(BOOL);

private:
	long	m_lRef;
};

#endif	//SERVER_MATH_H__