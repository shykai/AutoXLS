#include <map>

#include "PublicDef.h"

class MatchExcel_V2
{
public:

	MatchExcel_V2()
		:maker(wb.GetFormulaFactory())
		, offsetHeight(5)
	{
		// 		textFmt = wb.xformat();
		// 		titleRange->fontbold(BOLDNESS_BOLD);
		// 		titleRange->fillstyle(FILL_SOLID);
		// 		titleRange->fillfgcolor(CLR_GRAY40);
		// 		titleRange->halign(HALIGN_CENTER);
		// 		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		// 		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		// 		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		// 		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);


	};
	~MatchExcel_V2() {};

#include "MatchFormatter.h"


	void doExec(const MatchMap & inData)
	{
		std::list<struExcelAera> sumScoreList;
		std::list<struExcelAera> totalScoreList;

		uint32_t nowCol = 0;
		//Build Static Title and Stu Num&Name
		{
			uint32_t nowRow = 0;

			//Num&Name Title
			TextAera stuNoText(0, 0, 2, 0, L"学号");
			TextAera stuNameText(0, 1, 2, 1, L"姓名");

			TextAera lossText(3 + inData.stuCount , 0, 3 + inData.stuCount + 1, 1, L"失分");

			TextAera classNameText(rowFirstStaticsTitle + 0, 1, rowFirstStaticsTitle + 0, 1, L"分类");
			TextAera nodeTypeText(rowFirstStaticsTitle + 1, 1, rowFirstStaticsTitle + 1, 1, L"大题");
			TextAera nodeNameText(rowFirstStaticsTitle + 2, 1, rowFirstStaticsTitle + 2, 1, L"小题");
			TextAera nodeScoreText(rowFirstStaticsTitle + 3, 1, rowFirstStaticsTitle + 3, 1, L"单项总分");

			TextAera totalScoreText(rowFirstStaticsTitle + 5, 1, rowFirstStaticsTitle + 5, 1, L"应得分");
			TextAera finalScoreText(rowFirstStaticsTitle + 6, 1, rowFirstStaticsTitle + 6, 1, L"实得分");
			TextAera averageScoreText(rowFirstStaticsTitle + 7, 1, rowFirstStaticsTitle + 7, 1, L"得分率");


			nowRow += 3;

			InputArea numAndName(nowRow, nowCol, inData.stuCount, 2);
			nowCol += 2;
		}

		for (MatchClassLists::const_iterator classIter = inData.nodeList.begin(); classIter != inData.nodeList.end(); classIter++)
		{
			uint32_t nowRow = 0;
			TitleAera classTitle(nowRow, nowCol, classIter->matchCount + (classIter->matchType == Class_Deduct ? 1 : 0));
			classTitle.init(classIter->className, classIter->nodeLists, classIter->matchType);
			nowRow = classTitle.lastRow + 1;

			LossSumAera* lossSumPtr = NULL;
			LossTotalAera* lossTotalPtr = NULL;

			InputArea inputAera(nowRow, nowCol, inData.stuCount, classIter->matchCount);
			if (classIter->matchType == Class_Deduct)
			{
				lossSumPtr = new LossSumAera(inputAera);
				lossSumPtr->init(classIter->nodeLists);
				ScoreSumArea scoreAera(inputAera, rowTotalScore);
				lossTotalPtr = new LossTotalAera(*lossSumPtr);

				sumScoreList.push_back(scoreAera);
			}
			else if (classIter->matchType == Class_Add)
			{
				lossTotalPtr = new LossTotalAera(inputAera, rowTotalScore, inData.stuCount);
				sumScoreList.push_back(inputAera);
			}

			nowRow = rowFirstStaticsTitle;
			TitleAera scoreTitle(nowRow, nowCol, classIter->matchCount + (classIter->matchType == Class_Deduct ? 1 : 0));
			scoreTitle.init(classIter->className, classIter->nodeLists, classIter->matchType);
			nowRow = scoreTitle.lastRow + 1;

			InputArea scoreTotalAera(nowRow, nowCol, 1, classIter->matchCount);
			if (classIter->matchType == Class_Deduct)
			{
				ScoreSumArea scoreAera(scoreTotalAera, scoreTotalAera.firstRow, scoreTotalAera.lastCol + 1, scoreTotalAera.lastRow, scoreTotalAera.lastCol + 1);
				totalScoreList.push_back(scoreAera);
			}
			else
			{
				totalScoreList.push_back(scoreTotalAera);
			}

			if (classIter->matchType == Class_Deduct)
			{
				StaticsAera lossStatic(*lossSumPtr, rowTotalScore); lossStatic.init(inData.stuCount);
				StaticsAera lossTotalStatic(*lossTotalPtr, rowTotalScore); lossTotalStatic.init(inData.stuCount);
			}
			else if (classIter->matchType == Class_Add)
			{
				StaticsAera lossTotalStatic(*lossTotalPtr, rowTotalScore); lossTotalStatic.init(inData.stuCount);
			}

			if (lossSumPtr)
			{
				delete lossSumPtr;
			}
			if (lossTotalPtr)
			{
				delete lossTotalPtr;
			}

			nowCol += classIter->matchCount;
			nowCol += classIter->matchType == Class_Deduct ? 1 : 0;
		}

		if (inData.isSum)
		{
			std::wstring sumName;
			{
				bool isFirstName = true;
				for (MatchClassLists::const_iterator classIter = inData.nodeList.begin(); classIter != inData.nodeList.end(); classIter++)
				{
					if (isFirstName)
					{
						sumName = classIter->className;
						isFirstName = false;
					}
					else
					{
						sumName += L"+";
						sumName += classIter->className;
					}
				}
			}

			uint32_t nowRow = 0;
			TextAera sumTitle(nowRow, nowCol, nowRow + 2, nowCol, sumName);

			ScoreSumArea sumScore(sumScoreList);
			LossTotalAera sumLoss(sumScore, rowTotalScore, inData.stuCount);

			nowRow = rowFirstStaticsTitle;
			TextAera sumTitle2(nowRow, nowCol, nowRow + 2, nowCol, sumName);

			ScoreSumArea totalScore(totalScoreList);

			StaticsAera totalStatic(sumLoss, rowTotalScore); totalStatic.init(inData.stuCount);
		}
	}

	void doHeightAndWidth(const MatchMap & inData, worksheet* ws)
	{
		ws->rowheight(0, 20 * 20);
		ws->rowheight(1, 20 * 20);
		ws->rowheight(2, 20 * 20);
		for (uint32_t i = 3; i <= 3 + inData.stuCount; i++)
		{
			ws->rowheight(i, 18 * 20);
		}

		for (uint32_t i = rowFirstStaticsTitle; i < rowFirstStaticsScore + 3; i++)
		{
			ws->rowheight(i, 18 * 20);
		}

		ws->colwidth(0, 4 * 256);
		ws->colwidth(1, 12 * 256);
		uint32_t tmpCol = 2;
		for (MatchClassLists::const_iterator classIter = inData.nodeList.begin(); classIter != inData.nodeList.end(); classIter++)
		{
			MatchNodes tmpNodes = classIter->nodeLists;
			while (tmpNodes.size() > 0)
			{
				ws->colwidth(tmpCol, 8 * 256);
				tmpCol++;

				if (tmpNodes.front().nodeCount > 0)
				{
					tmpNodes.front().nodeCount--;
				}
				if (tmpNodes.front().nodeCount == 0)
				{
					tmpNodes.pop_front();
				}
			}

			if (classIter->matchType == Class_Deduct)
			{
				ws->colwidth(tmpCol, 8 * 256);
				tmpCol++;
			}
		}


	}

	void inputExcel(const MatchMap & inData)
	{
		worksheet* ws = wb.sheet(L"小分表");

		ws->defaultColwidth(8);
		ws->defaultRowHeight(18);


		wb.setColor(196, 215, 155, 9); //title
		wb.setColor(250, 191, 143, 10); //func
		wb.setColor(184, 204, 228, 11); //stu

		rowFirstStaticsTitle = 3 + inData.stuCount + 2 + offsetHeight; //TitleHeight=3, StudentCount, LossHeight=2, OffsetHeight
		rowTotalScore = rowFirstStaticsTitle + 3;
		rowFirstStaticsScore = rowTotalScore + 2;

		AeraManager::instance()->init(&wb, ws);

		doExec(inData);

		doHeightAndWidth(inData, ws);

		AeraManager::instance()->uinit();
	}



	bool outputExcel(const std::string &outFilePath)
	{
		int err = wb.Dump(outFilePath);

		return err == 0;
	};

private:

private:
	workbook wb;
	expression_node_factory_t& maker;

	uint32_t rowFirstStaticsTitle;
	uint32_t rowTotalScore;

	uint32_t rowFirstStaticsScore;
	uint32_t offsetHeight;
};


class TestCase
{
public:
	TestCase() { test(); };

	void test()
	{
		MatchExcel_V2 newExcel;

		MatchMap inData;
		inData.stuCount = 35;
		inData.isSum = true;

		{
			MatchNodes firstClassNodes;
			firstClassNodes.push_back(MatchNode(L"1", 1));
			firstClassNodes.push_back(MatchNode(L"2", 1));
			firstClassNodes.push_back(MatchNode(L"3", 1));
			firstClassNodes.push_back(MatchNode(L"4", 1));
			firstClassNodes.push_back(MatchNode(L"5", 10));
			firstClassNodes.push_back(MatchNode(L"6", 5));
			firstClassNodes.push_back(MatchNode(L"7", 1));
			firstClassNodes.push_back(MatchNode(L"8", 1));
			firstClassNodes.push_back(MatchNode(L"9", 5));

			MatchClass firstClass(L"A", firstClassNodes);
			inData.nodeList.push_back(firstClass);
		}

		{
			MatchNodes firstClassNodes;
			firstClassNodes.push_back(MatchNode(L"B", 1));

			MatchClass firstClass(L"B");
			inData.nodeList.push_back(firstClass);
		}

		newExcel.inputExcel(inData);
		newExcel.outputExcel("test.xls");
	};
};
// static TestCase testOnde;
