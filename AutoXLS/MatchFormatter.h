#pragma once

#include "xlslib.h"
#include "common/xlstypes.h"

#include "PublicDef.h"

class AeraManager
{
public:
	AeraManager()
	{
	};
	~AeraManager() {};

	static AeraManager* instance()
	{
		static AeraManager inst;
		return &inst;
	}

	void init(workbook* _wb, worksheet* _ws)
	{
		wb = _wb;
		ws = _ws;
	}

	void uinit()
	{
		wb = NULL;
		ws = NULL;
	}

	void format(uint32_t row1, uint32_t col1,
		uint32_t row2, uint32_t col2, char*firstCellPos, char* lastCellPos)
	{
		snprintf(firstCellPos, 16, "%c%c%d", col1 / 26 == 0 ? ' ' : 'A' + col1 / 26 - 1, 'A' + col1 % 26, row1 + 1);

		snprintf(lastCellPos, 16, "%c%c%d", col2 / 26 == 0 ? ' ' : 'A' + col2 / 26 - 1, 'A' + col2 % 26, row2 + 1);
	}

	worksheet* getWS() { return ws; };
	expression_node_factory_t& getMaker() { return wb->GetFormulaFactory(); };
	workbook* getWB() { return wb; };

	expression_node_t * buildFuncSum( /*uint32_t target_row, uint32_t target_col,*/
		uint32_t first_row, uint32_t first_col, uint32_t last_row, uint32_t last_col)
	{
		cell_t* lefttop = ws->FindCellOrMakeBlank(first_row, first_col);
		cell_t* rightbottom = ws->FindCellOrMakeBlank(last_row, last_col);

		expression_node_t *area = getMaker().area(*lefttop, *rightbottom, CELL_RELATIVE_A1, CELLOP_AS_REFER);
		expression_node_t *areas[1];
		areas[0] = area;
		expression_node_t *f = getMaker().f(FUNC_SUM, 1, areas, CELL_DEFAULT);
		// 		ws->formula(target_row, target_col, f, true);

		return f;
	};

private:
	workbook *wb;
	worksheet* ws;
};


struct struExcelAera
{
	struExcelAera()
	{

	}

	struExcelAera(uint32_t row1, uint32_t col1,
		uint32_t row2, uint32_t col2)
	{
		firstRow = row1;
		firstCol = col1;

		lastRow = row2;
		lastCol = col2;
	}


	uint32_t firstRow, firstCol;
	uint32_t lastRow, lastCol;

};

class ExcelArea : public struExcelAera
{
public:
	ExcelArea()
	{

	}
	ExcelArea(uint32_t row1, uint32_t col1,
		uint32_t row2, uint32_t col2)
	{
		firstRow = row1;
		firstCol = col1;

		lastRow = row2;
		lastCol = col2;

		AeraManager::instance()->format(row1, col1, row2, col2, strFirstCell, strLastCell);

	}

	virtual ~ExcelArea()
	{
	}

	virtual void build(worksheet* ws, expression_node_factory_t& maker, workbook* wb) = 0;
	virtual void act(worksheet* ws)  = 0;


	char strFirstCell[16] = { 0 };
	char strLastCell[16] = { 0 };

};

class TextAera : public ExcelArea
{
public:
	TextAera(uint32_t row1, uint32_t col1,
		uint32_t row2, uint32_t col2, const std::wstring& _textString)
		:ExcelArea(row1, col1, row2, col2)
		, textString(_textString)
	{

	}

	~TextAera()
	{
		build(AeraManager::instance()->getWS(), AeraManager::instance()->getMaker(), AeraManager::instance()->getWB());
		act(AeraManager::instance()->getWS());
	}

	virtual void build(worksheet* ws, expression_node_factory_t& maker, workbook* wb) 
	{
		ws->merge(firstRow, firstCol, lastRow, lastCol);
		ws->label(firstRow, firstCol, textString);
	};
	virtual void act(worksheet* ws)
	{
		range* titleRange = ws->rangegroup(firstRow, firstCol, lastRow, lastCol);

		titleRange->fontbold(BOLDNESS_BOLD);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->fillfgcolor((color_name_t)40);
		titleRange->halign(HALIGN_CENTER);
		titleRange->valign(VALIGN_CENTER);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);
		titleRange->wrap(true);

		titleRange->locked(true);
	}

private:
	std::wstring textString;
};

//标题区域
class TitleAera : public ExcelArea
{
public:
	TitleAera(uint32_t row1, uint32_t col1, uint32_t colCount)
		:ExcelArea(row1, col1, row1 + 2, col1 + colCount - 1)
	{

	}

	~TitleAera()
	{
		build(AeraManager::instance()->getWS(), AeraManager::instance()->getMaker(), AeraManager::instance()->getWB());
		act(AeraManager::instance()->getWS());
	}

	void init(const std::wstring& calssName, const MatchNodes& nodeList, const MatchClassType& classType)
	{
		_className = calssName;
		_nodeLists = nodeList;
		_classType = classType;
	}

	virtual void build(worksheet* ws, expression_node_factory_t& maker, workbook* wb)
	{
		ws->merge(firstRow, firstCol, firstRow, lastCol);
		ws->label(firstRow, firstCol, _className);

		uint32_t curRow = firstRow + 1;
		uint32_t curCol = firstCol;

		for (MatchNodes::const_iterator iter = _nodeLists.begin(); iter != _nodeLists.end(); iter++)
		{
			if (iter->nodeCount > 1)
			{
				ws->merge(curRow, curCol, curRow, curCol + iter->nodeCount - 1);
				ws->label(curRow, curCol, iter->nodeName);

				for (uint32_t i = 0; i < iter->nodeCount; i++)
				{
					ws->label(curRow + 1, curCol + i, toString(i + 1));
				}
			}
			else
			{
				ws->merge(curRow, curCol, curRow + 1, curCol);
				ws->label(curRow, curCol, iter->nodeName);
			}
			curCol += iter->nodeCount;
		}

		if (_classType == Class_Deduct)
		{
			ws->merge(firstRow + 1, lastCol, lastRow, lastCol);
			ws->label(firstRow + 1, lastCol, _className);
		}
	}
	virtual void act(worksheet* ws)
	{
		range* titleRange = ws->rangegroup(firstRow, firstCol, lastRow, lastCol);

		titleRange->fontbold(BOLDNESS_BOLD);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->fillfgcolor((color_name_t)40);
		titleRange->halign(HALIGN_CENTER);
		titleRange->valign(VALIGN_CENTER);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);
		titleRange->wrap(true);

		titleRange->locked(true);

	}
private:
	MatchNodes _nodeLists;
	std::wstring _className;
	MatchClassType _classType;
};

//输入区域
class InputArea : public ExcelArea
{
public:
	InputArea(uint32_t row1, uint32_t col1, uint32_t rowCount, uint32_t ColCount)
		:ExcelArea(row1, col1, row1 + rowCount-1, col1 + ColCount-1)
	{

	}

	~InputArea()
	{
		build(AeraManager::instance()->getWS(), AeraManager::instance()->getMaker(), AeraManager::instance()->getWB());
		act(AeraManager::instance()->getWS());
	}

	virtual void build(worksheet* ws, expression_node_factory_t& maker, workbook* wb) {};
	virtual void act(worksheet* ws)
	{
		bool isBule = true;

		for (unsigned32_t row = firstRow; row <= lastRow; row++)
		{
			range* titleRange = ws->rangegroup(row, firstCol, row, lastCol);
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

};

//统计得分列
class ScoreSumArea : public ExcelArea
{
public:
	ScoreSumArea(const InputArea &_input, uint32_t row1, uint32_t col1,
		uint32_t row2, uint32_t col2)
		:ExcelArea(row1, col1, row2, col2)
	{
		expression_node_t * sumScore = AeraManager::instance()->buildFuncSum(_input.firstRow, _input.firstCol, _input.lastRow, _input.lastCol);
		AeraManager::instance()->getWS()->formula(firstRow, firstCol, sumScore, true);
	}

	ScoreSumArea(const InputArea &_input, uint32_t _rowTotalScord)
		:ExcelArea(_input.firstRow, _input.lastCol + 1, _input.lastRow, _input.lastCol + 1)
	{
		expression_node_factory_t& maker = AeraManager::instance()->getMaker();
		cell_t* totalScore = AeraManager::instance()->getWS()->FindCellOrMakeBlank(_rowTotalScord, firstCol);

		for (uint32_t curRow = firstRow; curRow <= lastRow; curRow++)
		{
			expression_node_t * sumLoss = AeraManager::instance()->buildFuncSum(curRow, _input.firstCol, curRow, _input.lastCol);
			expression_node_t * score = AeraManager::instance()->getMaker().op(OP_SUB, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), sumLoss);
			AeraManager::instance()->getWS()->formula(curRow, firstCol, score, true);
		}
	}

	ScoreSumArea(const std::list<struExcelAera> scoreList)
		:ExcelArea(scoreList.back().firstRow, scoreList.back().lastCol + 1, scoreList.back().lastRow, scoreList.back().lastCol + 1)
	{
		expression_node_factory_t& maker = AeraManager::instance()->getMaker();
		
		for (uint32_t i = firstRow; i <= lastRow; i++)
		{
			std::list<struExcelAera> tmpList = scoreList;

			cell_t * totalScore = AeraManager::instance()->getWS()->FindCellOrMakeBlank(i, tmpList.front().firstCol);
			tmpList.pop_front();

			cell_t* plusScore = AeraManager::instance()->getWS()->FindCellOrMakeBlank(i, tmpList.front().firstCol);
			tmpList.pop_front();

			expression_node_t * score = maker.op(OP_ADD, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*plusScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

			for (std::list<struExcelAera>::const_iterator iter = tmpList.begin(); iter != tmpList.end(); iter++)
			{
				cell_t* plusScore = AeraManager::instance()->getWS()->FindCellOrMakeBlank(i, iter->firstCol);
				score = maker.op(OP_ADD, score, maker.cell(*plusScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
			}
			AeraManager::instance()->getWS()->formula(i, firstCol, score, true);
		}

	}

	~ScoreSumArea()
	{
		build(AeraManager::instance()->getWS(), AeraManager::instance()->getMaker(), AeraManager::instance()->getWB());
		act(AeraManager::instance()->getWS());
	}

	virtual void build(worksheet* ws, expression_node_factory_t& maker, workbook* wb)
	{
		//TODO
	}

	virtual void act(worksheet* ws)
	{
		range* titleRange = ws->rangegroup(firstRow, firstCol, lastRow, lastCol);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->fillfgcolor((color_name_t)17);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);

		titleRange->locked(true);
	}

};

//丢分统计行
class LossSumAera : public ExcelArea
{
public:
	LossSumAera(const InputArea &_input)
		:ExcelArea(_input.lastRow +1, _input.firstCol, _input.lastRow+2, _input.lastCol)
		, sumFirstRow(_input.firstRow)
		, sumLastRow(_input.lastRow)
	{

	}

	void init(const MatchNodes& nodeList)
	{
		_nodeLists = nodeList;
	}

	~LossSumAera()
	{
		build(AeraManager::instance()->getWS(), AeraManager::instance()->getMaker(), AeraManager::instance()->getWB());
		act(AeraManager::instance()->getWS());
	}

	virtual void build(worksheet* ws, expression_node_factory_t& maker, workbook* wb)
	{
		uint32_t curRow = firstRow;
		uint32_t curCol = firstCol;

		for (MatchNodes::const_iterator iter = _nodeLists.begin(); iter != _nodeLists.end(); iter++)
		{
			if (iter->nodeCount > 1)
			{
				for (uint32_t i = 0; i < iter->nodeCount; i++)
				{
					expression_node_t * f = AeraManager::instance()->buildFuncSum(sumFirstRow, curCol + i, sumLastRow, curCol + i);
					ws->formula(curRow, curCol + i, f, true);
				}

				ws->merge(curRow + 1, curCol, curRow + 1, curCol + iter->nodeCount - 1);


				expression_node_t * f = AeraManager::instance()->buildFuncSum( curRow, curCol, curRow, curCol + iter->nodeCount - 1);
				ws->formula(curRow + 1, curCol, f, true);

			}
			else
			{
				ws->merge(curRow, curCol, curRow + 1, curCol);
				expression_node_t * f = AeraManager::instance()->buildFuncSum(sumFirstRow, curCol, sumLastRow, curCol);
				ws->formula(curRow, curCol, f, true);
			}
			curCol += iter->nodeCount;
		}
	}

	virtual void act(worksheet* ws)
	{
		range* titleRange = ws->rangegroup(firstRow, firstCol, lastRow, lastCol);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->halign(HALIGN_CENTER);
		titleRange->valign(VALIGN_CENTER);
		titleRange->fillfgcolor((color_name_t)17);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);

		titleRange->locked(true);
	}

private:
	MatchNodes _nodeLists;
	uint32_t sumFirstRow, sumLastRow;
};

//总丢分总和计算
class LossTotalAera : public ExcelArea
{
public:
	LossTotalAera (const LossSumAera& _lossSum)
		:ExcelArea(_lossSum.firstRow, _lossSum.lastCol + 1, _lossSum.lastRow, _lossSum.lastCol+1)
	{
		AeraManager::instance()->getWS()->merge(firstRow, firstCol, lastRow, lastCol);
		expression_node_t * losFunc = AeraManager::instance()->buildFuncSum( /*curRow, curCol,*/ _lossSum.firstRow, _lossSum.firstCol, _lossSum.firstRow, _lossSum.lastCol);
		AeraManager::instance()->getWS()->formula(firstRow, firstCol, losFunc, true);

	}

	LossTotalAera(const InputArea& _input, uint32_t rowTotalScore, uint32_t stuCount)
		:ExcelArea(_input.lastRow + 1, _input.firstCol, _input.lastRow + 2, _input.lastCol )
	{
		expression_node_factory_t& maker = AeraManager::instance()->getMaker();
		worksheet* ws = AeraManager::instance()->getWS();

		ws->merge(firstRow, firstCol, lastRow, lastCol);

		cell_t *oneTotal = ws->FindCellOrMakeBlank(rowTotalScore, firstCol);
		expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

		expression_node_t * scoreFunc = AeraManager::instance()->buildFuncSum(_input.firstRow, _input.firstCol, _input.lastRow, _input.lastCol);

		expression_node_t * losFunc = maker.op(xlslib_core::OP_SUB, totalFunc, scoreFunc);

		ws->formula(firstRow, firstCol, losFunc, true);
	}

	LossTotalAera(const ScoreSumArea& _scoreSum, uint32_t rowTotalScore, uint32_t stuCount)
		:ExcelArea(_scoreSum.lastRow + 1, _scoreSum.firstCol, _scoreSum.lastRow + 2, _scoreSum.lastCol)
	{
		expression_node_factory_t& maker = AeraManager::instance()->getMaker();
		worksheet* ws = AeraManager::instance()->getWS();

		ws->merge(firstRow, firstCol, lastRow, lastCol);

		cell_t *oneTotal = ws->FindCellOrMakeBlank(rowTotalScore, firstCol);
		expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));
		
		expression_node_t * scoreFunc = AeraManager::instance()->buildFuncSum(_scoreSum.firstRow, _scoreSum.firstCol, _scoreSum.lastRow, _scoreSum.lastCol);

		expression_node_t * losFunc = maker.op(xlslib_core::OP_SUB, totalFunc, scoreFunc);

		ws->formula(firstRow, firstCol, losFunc, true);
	}

	~LossTotalAera()
	{
		build(AeraManager::instance()->getWS(), AeraManager::instance()->getMaker(), AeraManager::instance()->getWB());
		act(AeraManager::instance()->getWS());
	}

	virtual void build(worksheet* ws, expression_node_factory_t& maker, workbook* wb)
	{

	}

	virtual void act(worksheet* ws)
	{
		range* titleRange = ws->rangegroup(firstRow, firstCol, lastRow, lastCol);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->fillfgcolor((color_name_t)17);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);

		titleRange->locked(true);
	}

};

class StaticsAera : public ExcelArea
{
public:
	StaticsAera(const LossSumAera& _lostSum, uint32_t rowTotalScore)
		:ExcelArea(rowTotalScore +2, _lostSum.firstCol, rowTotalScore + 4, _lostSum.lastCol)
		,_rowTotalScore(rowTotalScore)
		, _rowLossScore(_lostSum.firstRow)
		, staticsType(Statics_ScorePercent)
	{

	}

	StaticsAera(const LossTotalAera& _lostTotal, uint32_t rowTotalScore)
		:ExcelArea(rowTotalScore + 2, _lostTotal.firstCol, rowTotalScore + 4, _lostTotal.lastCol)
		,_rowTotalScore(rowTotalScore)
		, _rowLossScore(_lostTotal.firstRow)
		, staticsType(Statics_Average)
	{

	}

	~StaticsAera()
	{
		build(AeraManager::instance()->getWS(), AeraManager::instance()->getMaker(), AeraManager::instance()->getWB());
		act(AeraManager::instance()->getWS());
	}

	void init(uint32_t stuCount)
	{
		_stuCount = stuCount;
	}

	virtual void build(worksheet* ws, expression_node_factory_t& maker, workbook* wb)
	{
		uint32_t curRow = firstRow;

		for (uint32_t curCol = firstCol; curCol <= lastCol; curCol++)
		{
			//应得分
			cell_t *oneTotal = ws->FindCellOrMakeBlank(_rowTotalScore, curCol);
			expression_node_t *totalFunc = maker.op(xlslib_core::OP_MUL, maker.integer((signed32_t)_stuCount), maker.cell(*oneTotal, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

			ws->formula(curRow, curCol, totalFunc, true);

			//实得分
			cell_t *totalScore = ws->FindCellOrMakeBlank(curRow, curCol);
			cell_t *totalLoss = ws->FindCellOrMakeBlank(_rowLossScore, curCol);
			expression_node_t *actScore = maker.op(xlslib_core::OP_SUB, maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalLoss, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

			ws->formula(curRow + 1, curCol, actScore, true);

			if (staticsType == Statics_ScorePercent)
			{
				//得分率
				cell_t *realScore = ws->FindCellOrMakeBlank(curRow + 1, curCol);
				expression_node_t *scorePercent = maker.op(xlslib_core::OP_DIV, maker.cell(*realScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.cell(*totalScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE));

				xf_t* sxf1 = wb->xformat();
				sxf1->SetFormat(FMT_PERCENT2);
				ws->formula(curRow + 2, curCol, scorePercent, true, sxf1);
			}
			else if (staticsType == Statics_Average)
			{
				//平均分
				cell_t *realScore = ws->FindCellOrMakeBlank(curRow + 1, curCol);
				expression_node_t *scorePercent = maker.op(xlslib_core::OP_DIV, maker.cell(*realScore, CELL_RELATIVE_A1, CELLOP_AS_VALUE), maker.integer((signed32_t)_stuCount));

				xf_t* sxf1 = wb->xformat();
				sxf1->SetFormat(FMT_NUMBER2);
				ws->formula(curRow + 2, curCol, scorePercent, true, sxf1);
			}
		}
	}

	virtual void act(worksheet* ws)
	{
		range* titleRange = ws->rangegroup(firstRow, firstCol, lastRow, lastCol);
		titleRange->fillstyle(FILL_SOLID);
		titleRange->fillfgcolor((color_name_t)17);
		titleRange->borderstyle(BORDER_BOTTOM, BORDER_THIN);
		titleRange->borderstyle(BORDER_TOP, BORDER_THIN);
		titleRange->borderstyle(BORDER_LEFT, BORDER_THIN);
		titleRange->borderstyle(BORDER_RIGHT, BORDER_THIN);

		titleRange->locked(true);
	}

private:
	uint32_t _rowTotalScore;
	uint32_t _rowLossScore;
	uint32_t _stuCount;

	enum StaticsType
	{
		Statics_ScorePercent = 0,
		Statics_Average,
	}staticsType;
};