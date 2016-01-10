#ifndef SERVER_IMATH_H__
#define SERVER_IMATH_H__
#include <Windows.h>
#include <ShlObj.h>

// {0437734F-22C8-40C0-85BA-26839ABD9F0D}
DEFINE_GUID(CLSID_Math, 
	0x437734f, 0x22c8, 0x40c0, 0x85, 0xba, 0x26, 0x83, 0x9a, 0xbd, 0x9f, 0xd);

// {29550C42-E8B5-432D-A677-E56703A9CDD0}
DEFINE_GUID(IID_IMath, 
	0x29550c42, 0xe8b5, 0x432d, 0xa6, 0x77, 0xe5, 0x67, 0x3, 0xa9, 0xcd, 0xd0);

// {99B2412C-BCAA-469A-856E-7E17B6999C13}
DEFINE_GUID(IID_IAdvandancedMath, 
	0x99b2412c, 0xbcaa, 0x469a, 0x85, 0x6e, 0x7e, 0x17, 0xb6, 0x99, 0x9c, 0x13);


class IMath : public IUnknown
{
public:

	STDMETHOD(Add)(long, long, long&)		PURE;
	STDMETHOD(Subtract)(long, long, long&)	PURE;
};

class IAdvancedMath : public IUnknown
{
	STDMETHOD(Factorial)(short, long&)		PURE;
	STDMETHOD(Fibonacci)(short, long&)		PURE;
};




#endif	//SERVER_IMATH_H__