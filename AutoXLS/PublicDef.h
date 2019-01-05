#pragma once

#include "xlslib.h"
using namespace xlslib_core;

struct MatchNode
{
	MatchNode(std::wstring _nodeName)
		:nodeName(_nodeName)
		, nodeCount(1)
	{
	};

	MatchNode(std::wstring _nodeName, uint32_t _nodeCount)
		:nodeName(_nodeName)
		, nodeCount(_nodeCount)
	{
	};

	std::wstring nodeName; //��Ŀ����
	uint32_t nodeCount; //С����
};
typedef std::list<MatchNode> MatchNodes;

enum MatchClassType
{
	Class_Deduct = 0, //��ʧ��ͳ��
	Class_Add //���÷�ͳ��
};
struct MatchClass
{
	MatchClass(std::wstring _className, const MatchNodes &_nodeLists, MatchClassType _matchType = Class_Deduct)
		:className(_className)
		, nodeLists(_nodeLists)
		, matchType(_matchType)
		, matchCount(0)
	{
		for (MatchNodes::const_iterator iter = nodeLists.begin(); iter != nodeLists.end(); iter++)
		{
			matchCount += iter->nodeCount;
		}
	};

	MatchClass(std::wstring _className, MatchClassType _matchType = Class_Add)
		:className(_className)
		, matchType(_matchType)
		, matchCount(1)
	{
		nodeLists.push_back(MatchNode(className));
	};

	MatchNodes nodeLists;
	std::wstring className;
	MatchClassType matchType;
	uint32_t matchCount;
};


typedef std::list<MatchClass> MatchClassLists;

struct MatchMap
{
	MatchMap()
		:stuCount(1)
		, isSum(false)
	{};

	uint32_t stuCount; //ѧ������
	MatchClassLists nodeList; //��Ŀ��Ϣ

	bool isSum;
};

static std::string toString(uint32_t valInt)
{
	char tmp[8] = { 0 };
	snprintf(tmp, 8, "%u", valInt);

	return std::string(tmp);
};

static std::string toColChar(uint32_t col)
{
	char tmp[2] = { 0 };
	tmp[0] = col;

	return std::string(tmp);
};


