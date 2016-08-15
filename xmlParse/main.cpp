#include <iostream>
#include <bitset>
#include <string>
#include "uimarkup.h"

using namespace std;

void readXmlFunc(XmlMarkup markUp)
{
	LPCTSTR filePath = L"C:\\Users\\weichong\\Desktop\\test.xml";
	LPCTSTR xmlContentTest = L"<note>test	\
							  <to time=\"20160812\">George</to> \
							  <from>John</from>	\
							  <heading>Reminder</heading>	\
							  <body>Don't forget the meeting!</body>	\
							  </note>";

	markUp.Load(xmlContentTest);

	XmlMarkupNode rootNode = markUp.GetRoot();
	wstring strName = rootNode.GetName();
	wstring strValue = rootNode.GetValue();

	XmlMarkupNode toNode = rootNode.GetChild(L"to");
	strName = toNode.GetName();
	strValue = toNode.GetValue();

	int attrCnt = toNode.GetAttributeCount();
	wstring timeAttr = toNode.GetAttributeName(0);
	wstring timeAttrValue = toNode.GetAttributeValue(L"time");
}

void writeXmlFunc(XmlMarkup &markUp)
{
	XmlMarkupNode rootNode;
	{
		char*rootName = "note";
		rootNode.SetName(rootName);
		char *rootValue = "value";
		rootNode.SetValue(rootValue);
		
		XmlAttribute attribute;
		char *attributeName ="size";
		attribute.SetName(attributeName);
		char *attributeValue = "100";
		attribute.SetValue(attributeValue);
		rootNode.AddAttribute(attribute);
	}
	
    XmlMarkupNode toNode;
    {
        toNode.SetName("today");
        toNode.SetValue("Sunday");
    }
	
	XmlMarkupNode remarkNode;
	{
        remarkNode.SetName("weather");
        remarkNode.SetValue("goo&&d day");

        XmlAttribute xmlAttr;
        xmlAttr.SetName("out");
        xmlAttr.SetValue("y\"es");
        remarkNode.AddAttribute(xmlAttr);
	}

    toNode.AddChild(remarkNode);
    rootNode.AddChild(toNode);
    markUp.SetRootNode(rootNode);
}

int main()
{
	XmlMarkup markUp;
	readXmlFunc(markUp);

	XmlMarkup writeMarkUp;
	writeXmlFunc(writeMarkUp);

    std::string ttStr = writeMarkUp.getXmlStr();
	system("puase");
}