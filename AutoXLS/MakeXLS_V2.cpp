#include <map>

#include "PublicDef.h"

class MatchExcel_V2
{
public:

	MatchExcel_V2()
		:maker(wb.GetFormulaFactory())
		, offsetLine(6)
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
		uint32_t rowFirstStatics = 10 + inData.stuCount;
		uint32_t rowTotalScore = rowFirstStatics + 3;
		//Build Static Title and Stu Num&Name
		{
			uint32_t nowRow = 0;

			//Num&Name Title
			nowRow += 3;

			InputArea numAndName(nowRow, nowCol, inData.stuCount, 2);
			nowCol += 2;
		}

		for (MatchClassLists::const_iterator classIter = inData.nodeList.begin(); classIter != inData.nodeList.end(); classIter++)
		{
			uint32_t nowRow = 0;
			TitleAera classTitle(nowRow, nowCol, classIter->matchCount+ (classIter->matchType == Class_Deduct?1:0));
			classTitle.init(classIter->className, classIter->nodeLists, classIter->matchType);
			nowRow = classTitle.lastRow + 1;

			LossSumAera* lossSumPtr = NULL;
			LossTotalAera* lossTotalPtr = NULL;

			InputArea inputAera(nowRow, nowCol, inData.stuCount, classIter->matchCount);
			if (classIter->matchType == Class_Deduct)
			{
				lossSumPtr = new LossSumAera(inputAera);
				ScoreSumArea scoreAera(inputAera, rowTotalScore);
				lossTotalPtr = new LossTotalAera(*lossSumPtr);

				sumScoreList.push_back(scoreAera);
			}
			else if (classIter->matchType == Class_Add)
			{
				lossTotalPtr = new LossTotalAera(inputAera);
				sumScoreList.push_back(inputAera);
			}

			nowRow = rowFirstStatics;
			TitleAera scoreTitle(nowRow, nowCol, classIter->matchCount + (classIter->matchType == Class_Deduct ? 1 : 0));
			classTitle.init(classIter->className, classIter->nodeLists, classIter->matchType);
			nowRow = scoreTitle.lastRow + 1;

			InputArea scoreTotalAera(nowRow, nowCol, 1, classIter->matchCount);
			if (classIter->matchType == Class_Deduct)
			{
				struExcelAera scoreAera(scoreTotalAera.firstRow, scoreTotalAera.lastCol + 1, scoreTotalAera.lastRow, scoreTotalAera.lastCol + 1);
// 				LossTotalAera scoreAera(scoreTotalAera); //?
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

			nowCol += classIter->matchCount;
			nowCol += classIter->matchType == Class_Deduct?1:0;
		}

		if (inData.isSum)
		{
			uint32_t nowRow = 0;
			TitleAera sumTitle(nowRow, nowCol, 1);

			ScoreSumArea sumScore(sumScoreList);
			LossTotalAera sumLoss(sumScore);

			nowRow = rowFirstStatics;
			TitleAera scoreTitle(nowRow, nowCol, 1);

			ScoreSumArea totalScore(totalScoreList);

			StaticsAera totalStatic(sumLoss, rowTotalScore);
		}
	}

	void inputExcel(const MatchMap & inData)
	{
		worksheet* ws = wb.sheet(L"Result");

		ws->defaultColwidth(8);
		ws->defaultRowHeight(18);

		wb.setColor(196, 215, 155, 9); //title
		wb.setColor(250, 191, 143, 10); //func
		wb.setColor(184, 204, 228, 11); //stu

		AeraManager::instance()->init(&wb, ws);

		doExec(inData);

		AeraManager::instance()->uinit();
	}

// 	void inputExcel(const MatchMap & inData)
// 	{
// 		uint32_t SumCol; //总分列
// 
// 		uint32_t lossRow; //失分行
// 
// 
// 		worksheet* ws = wb.sheet(L"统分表");
// 
// 		ws->defaultColwidth(8);
// 		ws->defaultRowHeight(18);
// 
// 		wb.setColor(196, 215, 155, 9); //title
// 		wb.setColor(250, 191, 143, 10); //func
// 		wb.setColor(184, 204, 228, 11); //stu
// 
// 
// 		uint32_t curCol = 0;
// 		uint32_t curRow = 0;
// 
// 		//学号
// 		ws->merge(curRow, curCol, curRow + 1, curCol);
// 		ws->label(curRow, curCol, L"学号");
// 		curCol++;
// 
// 		//姓名
// 		ws->merge(curRow, curCol, curRow + 1, curCol);
// 		ws->label(curRow, curCol, L"姓名");
// 		curCol++;
// 
// 		buildTitle(ws, curRow, curCol, inData.nodeList);
// 
// 		//总分
// 		SumCol = curCol;
// 		ws->merge(curRow, curCol, curRow + 1, curCol);
// 		ws->label(curRow, curCol, inData.totalTitle);
// 
// 		//附加题
// 		if (inData.isPlusNode)
// 		{
// 			curCol++;
// 			ws->merge(curRow, curCol, curRow + 1, curCol);
// 			ws->label(curRow, curCol, inData.plusTitle);
// 
// 			if (inData.isSum)
// 			{
// 				curCol++;
// 				ws->merge(curRow, curCol, curRow + 1, curCol);
// 				ws->label(curRow, curCol, inData.totalTitle + L"+" + inData.plusTitle);
// 			}
// 		}
// 
// 		actTitle(ws, curRow, 0, curRow + 1, curCol);
// 
// 		curRow += 2;
// 
// 		//姓名表
// 
// 		cell_t* totalScore = ws->FindCellOrMakeBlank(4 + offsetLine + inData.stuCount, SumCol);
// 		for (uint32_t i = curRow; i < curRow + inData.stuCount; i++)
// 		{
// 			expression_node_t * sumLoss = buildFuncSum(ws, i, 2, i, SumCol - 1);
// 			expression_node_t * score = maker.op(OP_SUB, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), sumLoss);
// 			ws->formula(i, SumCol, score, true);
// 		}
// 
// 		if (inData.isPlusNode && inData.isSum)
// 		{
// 			for (uint32_t i = curRow; i < curRow + inData.stuCount; i++)
// 			{
// 				cell_t * totalScore = ws->FindCellOrMakeBlank(i, SumCol);
// 				cell_t* plusScore = ws->FindCellOrMakeBlank(i, SumCol + 1);
// 				expression_node_t * score = maker.op(OP_ADD, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*plusScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
// 				ws->formula(i, SumCol + 2, score, true);
// 			}
// 		}
// 
// 		actStu(ws, curRow, 0, curRow + inData.stuCount - 1, curCol, SumCol, inData.isPlusNode && inData.isSum);
// 
// 		curRow += inData.stuCount;
// 
// 		//失分
// 		lossRow = curRow;
// 		curCol = 0;
// 		ws->merge(curRow, curCol, curRow + 1, curCol + 1);
// 		ws->label(curRow, curCol, L"失分");
// 		actTitle(ws, curRow, curCol, curRow + 1, curCol + 1);
// 
// 		curCol = 2;
// 
// 
// 		//失分统计
// 		buildLoss(ws, curRow, curCol, inData.nodeList, inData.stuCount);
// 
// 		//失分总分
// 		ws->merge(curRow, curCol, curRow + 1, curCol);
// 		expression_node_t * losFunc = buildFuncSum(ws, /*curRow, curCol,*/ curRow, 2, curRow, curCol - 1);
// 		ws->formula(curRow, curCol, losFunc, true);
// 
// 		actFunc(ws, curRow, 2, curRow + 1, curCol);
// 
// 		curCol += 1;
// 
// 		ws->rowheight(0, 20 * 20);
// 		ws->rowheight(1, 20 * 20);
// 		for (uint32_t i = 2; i <= curRow + 1; i++)
// 		{
// 			ws->rowheight(i, 18 * 20);
// 		}
// 		ws->colwidth(0, 4 * 256);
// 		ws->colwidth(1, 12 * 256);
// 
// 
// 		//小题单项分数
// 		{
// 			curRow += offsetLine;
// 			curCol = 1;
// 			ws->label(curRow, 1, L"大题");
// 			ws->label(curRow + 1, 1, L"小题");
// 			ws->label(curRow + 2, 1, L"单项总分");
// 
// 			curCol += 1;
// 			buildTitle(ws, curRow, curCol, inData.nodeList);
// 
// 			actEdit(ws, curRow + 2, 2, curRow + 2, curCol);
// 
// 			//试卷总分
// 			SumCol = curCol;
// 			ws->merge(curRow, curCol, curRow + 1, curCol);
// 			ws->label(curRow, curCol, inData.totalTitle);
// 
// 			//附加题总分
// 			if (inData.isPlusNode)
// 			{
// 				curCol++;
// 				ws->merge(curRow, curCol, curRow + 1, curCol);
// 				ws->label(curRow, curCol, inData.plusTitle);
// 
// 				actEdit(ws, curRow + 2, 2, curRow + 2, curCol);
// 
// 				if (inData.isSum)
// 				{
// 					curCol++;
// 					ws->merge(curRow, curCol, curRow + 1, curCol);
// 					ws->label(curRow, curCol, inData.totalTitle + L"+" + inData.plusTitle);
// 				}
// 			}
// 
// 			actTitle(ws, curRow, 1, curRow + 2, 1);
// 			actTitle(ws, curRow, 1, curRow + 1, curCol);
// 
// 
// 			//总分
// 			curRow += 2;
// 			curCol = SumCol;
// 			expression_node_t * totalFunc = buildFuncSum(ws, curRow, 2, curRow, curCol - 1);
// 			ws->formula(curRow, curCol, totalFunc, true);
// 
// 			actFunc(ws, curRow, curCol, curRow, curCol);
// 
// 			if (inData.isPlusNode && inData.isSum)
// 			{
// 				cell_t * totalScore = ws->FindCellOrMakeBlank(curRow, curCol);
// 				cell_t* plusScore = ws->FindCellOrMakeBlank(curRow, curCol + 1);
// 				expression_node_t * score = maker.op(OP_ADD, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*plusScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
// 				ws->formula(curRow, curCol + 2, score, true);
// 
// 				actFunc(ws, curRow, curCol + 2, curRow, curCol + 2);
// 			}
// 		}
// 
// 		//丢分统计
// 		{
// 			curRow += 2;
// 			curCol = 1;
// 			ws->label(curRow, 1, L"应得分");
// 			ws->label(curRow + 1, 1, L"实得分");
// 			ws->label(curRow + 2, 1, L"得分率");
// 
// 			actTitle(ws, curRow, 1, curRow + 2, 1);
// 
// 			curCol += 1;
// 			//单项
// 			for (MatchNodes::const_iterator iter = inData.nodeList.begin(); iter != inData.nodeList.end(); iter++)
// 			{
// 				for (uint32_t i = 0; i < iter->nodeCount; i++)
// 				{
// 					//应得分
// 					cell_t *oneTotal = ws->FindCellOrMakeBlank(curRow - 2, curCol + i);
// 					expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)inData.stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
// 
// 					ws->formula(curRow, curCol + i, totalFunc, true);
// 
// 					//实得分
// 					cell_t *totalScore = ws->FindCellOrMakeBlank(curRow, curCol + i);
// 					cell_t *totalLoss = ws->FindCellOrMakeBlank(lossRow, curCol + i);
// 					expression_node_t *actScore = maker.op(xlslib_core::OP_SUB, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalLoss, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
// 
// 					ws->formula(curRow + 1, curCol + i, actScore, true);
// 
// 					//得分率
// 					cell_t *realScore = ws->FindCellOrMakeBlank(curRow + 1, curCol + i);
// 					expression_node_t *scorePercent = maker.op(xlslib_core::OP_DIV, maker.cell(*realScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
// 
// 					xf_t* sxf1 = wb.xformat();
// 					sxf1->SetFormat(FMT_PERCENT2);
// 					ws->formula(curRow + 2, curCol + i, scorePercent, true, sxf1);
// 				}
// 				curCol += iter->nodeCount;
// 			}
// 
// 			//总分
// 			{
// 				//应得分
// 				cell_t *oneTotal = ws->FindCellOrMakeBlank(curRow - 2, curCol);
// 				expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)inData.stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
// 
// 				ws->formula(curRow, curCol, totalFunc, true);
// 
// 				//实得分
// 				cell_t *totalScore = ws->FindCellOrMakeBlank(curRow, curCol);
// 				cell_t *totalLoss = ws->FindCellOrMakeBlank(lossRow, curCol);
// 				expression_node_t *actScore = maker.op(xlslib_core::OP_SUB, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalLoss, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
// 
// 				ws->formula(curRow + 1, curCol, actScore, true);
// 
// 				//得分率
// 				cell_t *realScore = ws->FindCellOrMakeBlank(curRow + 1, curCol);
// 				expression_node_t *scorePercent = maker.op(xlslib_core::OP_DIV, maker.cell(*realScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.integer(inData.stuCount));
// 
// 				xf_t* sxf1 = wb.xformat();
// 				sxf1->SetFormat(FMT_NUMBER2);
// 				ws->formula(curRow + 2, curCol, scorePercent, true, sxf1);
// 
// 			}
// 
// 			//附加分
// 			if (inData.isPlusNode)
// 			{
// 				curCol += 1;
// 
// 				//应得分
// 				cell_t *oneTotal = ws->FindCellOrMakeBlank(curRow - 2, curCol);
// 				expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)inData.stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
// 
// 				ws->formula(curRow, curCol, totalFunc, true);
// 
// 				//实得分
// 				cell_t *totalScore = ws->FindCellOrMakeBlank(curRow, curCol);
// 				expression_node_t *actScore = buildFuncSum(ws, 2, curCol, 2 + inData.stuCount - 1, curCol);
// 
// 				ws->formula(curRow + 1, curCol, actScore, true);
// 
// 				//得分率
// 				cell_t *realScore = ws->FindCellOrMakeBlank(curRow + 1, curCol);
// 				expression_node_t *scorePercent = maker.op(xlslib_core::OP_DIV, maker.cell(*realScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.integer(inData.stuCount));
// 
// 				xf_t* sxf1 = wb.xformat();
// 				sxf1->SetFormat(FMT_NUMBER2);
// 				ws->formula(curRow + 2, curCol, scorePercent, true, sxf1);
// 
// 				if (inData.isSum)
// 				{
// 					curCol += 1;
// 
// 					//应得分
// 					cell_t *oneTotal = ws->FindCellOrMakeBlank(curRow - 2, curCol);
// 					expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)inData.stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
// 
// 					ws->formula(curRow, curCol, totalFunc, true);
// 
// 					//实得分
// 					cell_t *totalScore = ws->FindCellOrMakeBlank(curRow, curCol);
// 					expression_node_t *actScore = buildFuncSum(ws, 2, curCol, 2 + inData.stuCount - 1, curCol);
// 
// 					ws->formula(curRow + 1, curCol, actScore, true);
// 
// 					//得分率
// 					cell_t *realScore = ws->FindCellOrMakeBlank(curRow + 1, curCol);
// 					expression_node_t *scorePercent = maker.op(xlslib_core::OP_DIV, maker.cell(*realScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.integer(inData.stuCount));
// 
// 					xf_t* sxf1 = wb.xformat();
// 					sxf1->SetFormat(FMT_NUMBER2);
// 					ws->formula(curRow + 2, curCol, scorePercent, true, sxf1);
// 				}
// 			}
// 
// 			actFunc(ws, curRow, 2, curRow + 2, curCol);
// 
// 			for (uint32_t i = SumCol - 4; i < SumCol + 3; i++)
// 			{
// 				ws->rowheight(i, 18 * 20);
// 			}
// 
// 			MatchNodes tmpNodes = inData.nodeList;
// 			uint32_t tmpCol = 2;
// 			while (tmpNodes.size() > 0)
// 			{
// 				ws->colwidth(tmpCol, 8 * 256);
// 				tmpCol++;
// 
// 				if (tmpNodes.front().nodeCount > 0)
// 				{
// 					tmpNodes.front().nodeCount--;
// 				}
// 				if (tmpNodes.front().nodeCount == 0)
// 				{
// 					tmpNodes.pop_front();
// 				}
// 			}
// 		}
// 	}

	bool outputExcel(const std::string &outFilePath)
	{
		int err = wb.Dump(outFilePath);

		return err == 0;
	};

private:



private:

	expression_node_t * buildFuncSum(worksheet* ws, /*uint32_t target_row, uint32_t target_col,*/
		uint32_t first_row, uint32_t first_col, uint32_t last_row, uint32_t last_col)
	{
		cell_t* lefttop = ws->FindCellOrMakeBlank(first_row, first_col);
		cell_t* rightbottom = ws->FindCellOrMakeBlank(last_row, last_col);

		expression_node_t *area = maker.area(*lefttop, *rightbottom, CELL_RELATIVE_A1, CELLOP_AS_REFER);
		expression_node_t *areas[1];
		areas[0] = area;
		expression_node_t *f = maker.f(FUNC_SUM, 1, areas, CELL_DEFAULT);
// 		ws->formula(target_row, target_col, f, true);

		return f;
	};

// 	void buildTitle(worksheet* ws, uint32_t &curRow, uint32_t &curCol, const MatchNodes &nodeList)
// 	{
// 		//题目表
// 		for (MatchNodes::const_iterator iter = nodeList.begin(); iter != nodeList.end(); iter++)
// 		{
// 			if (iter->nodeCount > 1)
// 			{
// 				ws->merge(curRow, curCol, curRow, curCol + iter->nodeCount - 1);
// 				ws->label(curRow, curCol, iter->nodeName);
// 
// 				for (uint32_t i = 0; i < iter->nodeCount; i++)
// 				{
// 					ws->label(curRow + 1, curCol + i, toString(i + 1));
// 				}
// 			}
// 			else
// 			{
// 				ws->merge(curRow, curCol, curRow + 1, curCol);
// 				ws->label(curRow, curCol, iter->nodeName);
// 			}
// 			curCol += iter->nodeCount;
// 		}
// 	}
// 	void buildTitle(worksheet* ws, uint32_t &curRow, uint32_t &curCol, const MatchWithClasses &newNodeList)
// 	{
// 		for (MatchWithClasses::const_iterator iter = newNodeList.begin(); iter != newNodeList.end(); iter++)
// 		{
// 			buildTitle(ws, curRow, curCol, iter->second);
// 		}
// 	}


	void buildLoss(worksheet* ws, uint32_t &curRow, uint32_t &curCol, const MatchNodes &nodeList, uint32_t stuCount)
	{
		//题目表
		for (MatchNodes::const_iterator iter = nodeList.begin(); iter != nodeList.end(); iter++)
		{
			if (iter->nodeCount > 1)
			{
				for (uint32_t i = 0; i < iter->nodeCount; i++)
				{
					expression_node_t * f = buildFuncSum(ws, /*curRow, curCol + i,*/ 2, curCol + i, 2 + stuCount - 1, curCol + i);
					ws->formula(curRow, curCol + i, f, true);
				}

				ws->merge(curRow + 1, curCol, curRow + 1, curCol + iter->nodeCount - 1);


				expression_node_t * f = buildFuncSum(ws, /*curRow + 1, curCol,*/ curRow, curCol, curRow, curCol + iter->nodeCount - 1);
				ws->formula(curRow + 1, curCol, f, true);

			}
			else
			{
				ws->merge(curRow, curCol, curRow + 1, curCol);
				expression_node_t * f = buildFuncSum(ws, /*curRow, curCol,*/ 2, curCol, 2 + stuCount - 1, curCol);
				ws->formula(curRow, curCol, f, true);
			}
			curCol += iter->nodeCount;
		}
	}

	void actTitle(worksheet* ws, unsigned32_t row1, unsigned32_t col1,
		unsigned32_t row2, unsigned32_t col2)
	{
		range* titleRange = ws->rangegroup(row1, col1, row2, col2);

		titleRange->fontbold(BOLDNESS_BOLD);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->fillfgcolor((color_name_t)40);
		titleRange->halign(HALIGN_CENTER);
		titleRange->valign(VALIGN_CENTER);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);

		titleRange->locked(true);
	}

	void actFunc(worksheet* ws, unsigned32_t row1, unsigned32_t col1,
		unsigned32_t row2, unsigned32_t col2)
	{
		range* titleRange = ws->rangegroup(row1, col1, row2, col2);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->fillfgcolor((color_name_t)17);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);

		titleRange->locked(true);
	}

	void actStu(worksheet* ws, unsigned32_t row1, unsigned32_t col1,
		unsigned32_t row2, unsigned32_t col2, unsigned32_t sumColNo, bool isSumTotal)
	{
		bool isBule = true;

		for (unsigned32_t row = row1; row <= row2; row++)
		{
			range* titleRange = ws->rangegroup(row, col1, row, col2);
			if (isBule)
			{
				titleRange->fillstyle(FILL_SOLID);
				titleRange->fillfgcolor((color_name_t)28);
			}
			titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
			titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
			titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
			titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);
			titleRange->locked(false);

			isBule = !isBule;
		}

		range* stuNo = ws->rangegroup(row1, 0, row2, 0);
		stuNo->fillstyle(FILL_SOLID);
		stuNo->fillfgcolor((color_name_t)28);
		
		range* sumCol = ws->rangegroup(row1, sumColNo, row2, sumColNo);
		sumCol->locked(true);
		if (sumColNo == col2)
		{
			sumCol->fillstyle(FILL_SOLID);
			sumCol->fillfgcolor((color_name_t)28);
		}
		else
		{
			range* lastCol = ws->rangegroup(row1, col2, row2, col2);
			lastCol->fillstyle(FILL_SOLID);
			lastCol->fillfgcolor((color_name_t)28);
			if (isSumTotal)
			{
				lastCol->locked(true);
			}
		}
	}

	void actEdit(worksheet* ws, unsigned32_t row1, unsigned32_t col1,
		unsigned32_t row2, unsigned32_t col2)
	{
		bool isBule = true;

		for (unsigned32_t row = row1; row <= row2; row++)
		{
			range* titleRange = ws->rangegroup(row, col1, row, col2);
			if (isBule)
			{
				titleRange->fillstyle(FILL_SOLID);
				titleRange->fillfgcolor((color_name_t)28);
			}
			titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
			titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
			titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
			titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);
			titleRange->locked(false);

			isBule = !isBule;
		}
	}


private:
	workbook wb;
	expression_node_factory_t& maker;

	uint32_t offsetLine;
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
