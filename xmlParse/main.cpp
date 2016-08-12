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
	//	markUp.LoadFromFile(filePath);

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
	//todo�� ��Ż��������ʹ����ʽ��Ȼ������������ƴ����ӿ�
	XmlMarkupNode rootNode;
	{
		//todo:
		char*rootName = "note";
		rootNode.SetName(rootName);
		char *rootValue = "value";
		rootNode.SetValue(rootValue);
		
		XmlAttribute attribute;
		char *attributeName ="size";
		attribute.SetName(attributeName);
		char *attributeValue = "100";
		attribute.SetValue(attributeValue);
		rootNode.AddAttribute((void *)&attribute);
	}
	markUp.SetRootNode(rootNode);

	XmlMarkupNode toNode;
	{
		//todo:  ����node�ĸ�������
		//�������������ܸ�����
	}
	rootNode.AddChild((void*)&toNode);
	
	XmlMarkupNode remarkNode;
	{
		//todo: ����remarkNode��������
	}
	toNode.AddChild((void*)&remarkNode);

	XmlMarkupNode fromNode;
	{
		//todo: ����node�ĸ�������
	}
	rootNode.AddChild((void*)&toNode);
}

int main()
{
	XmlMarkup markUp;
	readXmlFunc(markUp);

	XmlMarkup writeMarkUp;
	writeXmlFunc(markUp);
	system("puase");
}