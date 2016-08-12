#ifndef __UIMARKUP_H__
#define __UIMARKUP_H__

#include <Windows.h>

enum
{
    XMLFILE_ENCODING_UTF8 = 0,
    XMLFILE_ENCODING_UNICODE = 1,
    XMLFILE_ENCODING_ASNI = 2,
};

class XmlMarkup;
class XmlMarkupNode;


class XmlMarkup
{
    friend class XmlMarkupNode;
public:
    XmlMarkup(LPCTSTR pstrXML = NULL);
    ~XmlMarkup();

	//load from pstr
    bool Load(LPCTSTR pstrXML);
    //load from file handle
	bool LoadFromMem(BYTE* pByte, DWORD dwSize, int encoding = XMLFILE_ENCODING_UTF8);
    bool LoadFromFile(LPCTSTR pstrFilename, int encoding = XMLFILE_ENCODING_UTF8);
    void Release();
    bool IsValid() const;

    void SetPreserveWhitespace(bool bPreserve = true);
    void GetLastErrorMessage(LPTSTR pstrMessage, SIZE_T cchMax) const;
    void GetLastErrorLocation(LPTSTR pstrSource, SIZE_T cchMax) const;

    XmlMarkupNode GetRoot();
	void SetRootNode(XmlMarkupNode rootNode);

private:
	//保存的是element在字符串中的位置，通过字符串指针首地址加上位置即可得到对应的字符串
    typedef struct tagXMLELEMENT
    {
        unsigned long iStart;			//element start location
        unsigned long iChild;			//
        unsigned long iNext;
        unsigned long iParent;
        unsigned long iData;
    } XMLELEMENT;

    LPTSTR m_pstrXML;					//xml字符串首地址
    XMLELEMENT* m_pElements;		
    unsigned long m_nElements;
    unsigned long m_nReservedElements;
    TCHAR m_szErrorMsg[100];
    TCHAR m_szErrorXML[50];
    bool m_bPreserveWhitespace;

private:
    bool _Parse();
    bool _Parse(LPTSTR& pstrText, ULONG iParent);
    XMLELEMENT* _ReserveElement();
    inline void _SkipWhitespace(LPTSTR& pstr) const;
    inline void _SkipWhitespace(LPCTSTR& pstr) const;
    inline void _SkipIdentifier(LPTSTR& pstr) const;
    inline void _SkipIdentifier(LPCTSTR& pstr) const;
    bool _ParseData(LPTSTR& pstrText, LPTSTR& pstrData, char cEnd);
    void _ParseMetaChar(LPTSTR& pstrText, LPTSTR& pstrDest);
    bool _ParseAttributes(LPTSTR& pstrText);
    bool _Failed(LPCTSTR pstrError, LPCTSTR pstrLocation = NULL);
};


class XmlMarkupNode
{
    friend class XmlMarkup;
public:
    XmlMarkupNode();
    XmlMarkupNode(XmlMarkup* pOwner, int iPos);

public:
    bool IsValid() const;

    XmlMarkupNode GetParent();
    XmlMarkupNode GetSibling();
    XmlMarkupNode GetChild();
    XmlMarkupNode GetChild(LPCTSTR pstrName);
	void AddChild(void*);

    bool HasSiblings() const;
    bool HasChildren() const;
    LPCTSTR GetName() const;
    LPCTSTR GetValue() const;
	void SetName(char* name);
	void SetValue(char* value);

    bool HasAttributes();
    bool HasAttribute(LPCTSTR pstrName);
    int GetAttributeCount();
    LPCTSTR GetAttributeName(int iIndex);
    LPCTSTR GetAttributeValue(int iIndex);
    LPCTSTR GetAttributeValue(LPCTSTR pstrName);
    bool GetAttributeValue(int iIndex, LPTSTR pstrValue, SIZE_T cchMax);
    bool GetAttributeValue(LPCTSTR pstrName, LPTSTR pstrValue, SIZE_T cchMax);

	void AddAttribute(void*);

private:
    void _MapAttributes();

    enum { MAX_XML_ATTRIBUTES = 64 };

    typedef struct
    {
        ULONG iName;
        ULONG iValue;
    } XMLATTRIBUTE;

    int m_iPos;
    int m_nAttributes;
    XMLATTRIBUTE m_aAttributes[MAX_XML_ATTRIBUTES];
    XmlMarkup* m_pOwner;
};


class XmlAttribute
{
public:
	void SetName(char*);
	void SetValue(char*);
};

#endif // __UIMARKUP_H__
