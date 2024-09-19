using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LIBEXCELMANIPULATOR
{
    public class XLBLUEPRINT
    {
        protected XLWorkbook? _XLblueprint = new XLWorkbook();
        protected IXLWorksheet? _mastering;
        protected IXLWorksheet? _realtime;

        protected IXLRange? _rangeMasterModelTable;
        protected IXLRange? _rangeMasterStep1Param;
        protected IXLRange? _rangeMasterStep2345Param;
        protected IXLRange? _rangeMasterDataHeader;
        protected IXLRange? _rangeMasterStep2;
        protected IXLRange? _rangeMasterStep3;
        //IXLRange _rangeMasterStep4;
        //IXLRange _rangeMasterStep5;

        protected IXLRange? _rangeRealtimeModelTable;
        protected IXLRange? _rangeRealtimeStep1Param;
        protected IXLRange? _rangeRealtimeStep2345Param;
        protected IXLRange? _rangeRealtimeDataHeader;
        protected IXLRange? _rangeRealtimeJudgement;
        protected IXLRange? _rangeRealtimeStep2;
        protected IXLRange? _rangeRealtimeStep3;
        //IXLRange _rangeRealtimeStep4;
        //IXLRange _rangeRealtimeStep5;

        protected List<IXLCell>? _cellMasterModelTableVarMap = new List<IXLCell>();
        protected IXLCell? _cellMasterModelName;
        protected IXLCell? _cellMasterDay;
        protected IXLCell? _cellMasterMonth;
        protected IXLCell? _cellMasterYear;
        protected IXLCell? _cellMasterHour;
        protected IXLCell? _cellMasterMinute;
        protected IXLCell? _cellMasterSecond;

        void _initMasterModelTableVarMap()
        {
            _cellMasterModelName = _rangeMasterModelTable.Cell(1, 2);
            _cellMasterDay = _rangeMasterModelTable.Cell(2, 2);
            _cellMasterMonth = _rangeMasterModelTable.Cell(2, 4);
            _cellMasterYear = _rangeMasterModelTable.Cell(2, 6);
            _cellMasterHour = _rangeMasterModelTable.Cell(3, 2);
            _cellMasterMinute = _rangeMasterModelTable.Cell(3, 4);
            _cellMasterSecond = _rangeMasterModelTable.Cell(3, 6);

            _cellMasterModelTableVarMap.Add(_cellMasterModelName);
            _cellMasterModelTableVarMap.Add(_cellMasterDay);
            _cellMasterModelTableVarMap.Add(_cellMasterMonth);
            _cellMasterModelTableVarMap.Add(_cellMasterYear);
            _cellMasterModelTableVarMap.Add(_cellMasterHour);
            _cellMasterModelTableVarMap.Add(_cellMasterMinute);
            _cellMasterModelTableVarMap.Add(_cellMasterSecond);
        }

        protected List<IXLCell>? _cellMasterStep1ParamVarMap = new List<IXLCell>();
        protected IXLCell? _cellMasterStep1Mode;
        protected IXLCell? _cellMasterStep1Stroke;
        protected IXLCell? _cellMasterStep1CompSpeed;
        protected IXLCell? _cellMasterStep1ExtnSpeed;
        protected IXLCell? _cellMasterStep1CycleCount;
        protected IXLCell? _cellMasterStep1MaxLoad;

        void _initMasterStep1ParamVarMap()
        {
            _cellMasterStep1Mode = _rangeMasterStep1Param.Cell(2, 4);
            _cellMasterStep1Stroke = _rangeMasterStep1Param.Cell(3, 4);
            _cellMasterStep1CompSpeed = _rangeMasterStep1Param.Cell(4, 4);
            _cellMasterStep1ExtnSpeed = _rangeMasterStep1Param.Cell(5, 4);
            _cellMasterStep1CycleCount = _rangeMasterStep1Param.Cell(6, 4);
            _cellMasterStep1MaxLoad = _rangeMasterStep1Param.Cell(7, 4);

            _cellMasterStep1ParamVarMap.Add(_cellMasterStep1Mode);
            _cellMasterStep1ParamVarMap.Add(_cellMasterStep1Stroke);
            _cellMasterStep1ParamVarMap.Add(_cellMasterStep1CompSpeed);
            _cellMasterStep1ParamVarMap.Add(_cellMasterStep1ExtnSpeed);
            _cellMasterStep1ParamVarMap.Add(_cellMasterStep1CycleCount);
            _cellMasterStep1ParamVarMap.Add(_cellMasterStep1MaxLoad);
        }

        protected List<IXLCell>? _cellMasterStep2345ParamVarMap = new List<IXLCell>();
        protected IXLCell? _cellMasterStep2Mode;
        protected IXLCell? _cellMasterStep2CompSpeed;
        protected IXLCell? _cellMasterStep2CompJudgePosMin;
        protected IXLCell? _cellMasterStep2CompJudgePosMax;
        protected IXLCell? _cellMasterStep2CompLoadRefPos;
        protected IXLCell? _cellMasterStep2ExtnSpeed;
        protected IXLCell? _cellMasterStep2ExtnJudgePosMin;
        protected IXLCell? _cellMasterStep2ExtnJudgePosMax;
        protected IXLCell? _cellMasterStep2ExtnLoadRefPos;
        protected IXLCell? _cellMasterStep2LoadRefTolerance;
        protected IXLCell? _cellMasterStep3Mode;
        protected IXLCell? _cellMasterStep3CompSpeed;
        protected IXLCell? _cellMasterStep3CompJudgePosMin;
        protected IXLCell? _cellMasterStep3CompJudgePosMax;
        protected IXLCell? _cellMasterStep3CompLoadRefPos;
        protected IXLCell? _cellMasterStep3ExtnSpeed;
        protected IXLCell? _cellMasterStep3ExtnJudgePosMin;
        protected IXLCell? _cellMasterStep3ExtnJudgePosMax;
        protected IXLCell? _cellMasterStep3ExtnLoadRefPos;
        protected IXLCell? _cellMasterStep3LoadRefTolerance;

        void _initMasterStep2345ParamVarMap()
        {
            _cellMasterStep2Mode = _rangeMasterStep2345Param.Cell(2, 5);
            _cellMasterStep2CompSpeed = _rangeMasterStep2345Param.Cell(3, 5);
            _cellMasterStep2CompJudgePosMin = _rangeMasterStep2345Param.Cell(4, 5);
            _cellMasterStep2CompJudgePosMax = _rangeMasterStep2345Param.Cell(5, 5);
            _cellMasterStep2CompLoadRefPos = _rangeMasterStep2345Param.Cell(6, 5);
            _cellMasterStep2ExtnSpeed = _rangeMasterStep2345Param.Cell(7, 5);
            _cellMasterStep2ExtnJudgePosMin = _rangeMasterStep2345Param.Cell(8, 5);
            _cellMasterStep2ExtnJudgePosMax = _rangeMasterStep2345Param.Cell(9, 5);
            _cellMasterStep2ExtnLoadRefPos = _rangeMasterStep2345Param.Cell(10, 5);
            _cellMasterStep2LoadRefTolerance = _rangeMasterStep2345Param.Cell(11, 5);
            _cellMasterStep3Mode = _rangeMasterStep2345Param.Cell(2, 6);
            _cellMasterStep3CompSpeed = _rangeMasterStep2345Param.Cell(3, 6);
            _cellMasterStep3CompJudgePosMin = _rangeMasterStep2345Param.Cell(4, 6);
            _cellMasterStep3CompJudgePosMax = _rangeMasterStep2345Param.Cell(5, 6);
            _cellMasterStep3CompLoadRefPos = _rangeMasterStep2345Param.Cell(6, 6);
            _cellMasterStep3ExtnSpeed = _rangeMasterStep2345Param.Cell(7, 6);
            _cellMasterStep3ExtnJudgePosMin = _rangeMasterStep2345Param.Cell(8, 6);
            _cellMasterStep3ExtnJudgePosMax = _rangeMasterStep2345Param.Cell(9, 6);
            _cellMasterStep3ExtnLoadRefPos = _rangeMasterStep2345Param.Cell(10, 6);
            _cellMasterStep3LoadRefTolerance = _rangeMasterStep2345Param.Cell(11, 6);

            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2Mode);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2CompSpeed);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2CompJudgePosMin);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2CompJudgePosMax);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2CompLoadRefPos);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2ExtnSpeed);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2ExtnJudgePosMin);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2ExtnJudgePosMax);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2ExtnLoadRefPos);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep2LoadRefTolerance);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3Mode);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3CompSpeed);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3CompJudgePosMin);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3CompJudgePosMax);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3CompLoadRefPos);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3ExtnSpeed);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3ExtnJudgePosMin);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3ExtnJudgePosMax);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3ExtnLoadRefPos);
            _cellMasterStep2345ParamVarMap.Add(_cellMasterStep3LoadRefTolerance);
        }

        protected List<IXLRange>? _cellMasterStep2VarMap = new List<IXLRange>();
        protected IXLRange? _cellMasterStep2CompStroke;
        protected IXLRange? _cellMasterStep2CompLoad;
        protected IXLRange? _cellMasterStep2CompLoadLower;
        protected IXLRange? _cellMasterStep2CompLoadUpper;
        protected IXLRange? _cellMasterStep2ExtnStroke;
        protected IXLRange? _cellMasterStep2ExtnLoad;
        protected IXLRange? _cellMasterStep2ExtnLoadLower;
        protected IXLRange? _cellMasterStep2ExtnLoadUpper;
        protected IXLRange? _cellMasterStep2DiffStroke;
        protected IXLRange? _cellMasterStep2DiffLoad;
        protected IXLRange? _cellMasterStep2DiffLoadLower;
        protected IXLRange? _cellMasterStep2DiffLoadUpper;

        void _initMasterStep2VarMap()
        {
            _cellMasterStep2CompStroke = _rangeMasterStep2.Range(4, 1, 205, 1);
            _cellMasterStep2CompLoad = _rangeMasterStep2.Range(4, 2, 205, 2);
            _cellMasterStep2CompLoadLower = _rangeMasterStep2.Range(4, 3, 205, 3);
            _cellMasterStep2CompLoadUpper = _rangeMasterStep2.Range(4, 4, 205, 4);
            _cellMasterStep2ExtnStroke = _rangeMasterStep2.Range(4, 5, 205, 5);
            _cellMasterStep2ExtnLoad = _rangeMasterStep2.Range(4, 6, 205, 6);
            _cellMasterStep2ExtnLoadLower = _rangeMasterStep2.Range(4, 7, 205, 7);
            _cellMasterStep2ExtnLoadUpper = _rangeMasterStep2.Range(4, 8, 205, 8);
            _cellMasterStep2DiffStroke = _rangeMasterStep2.Range(4, 9, 205, 9);
            _cellMasterStep2DiffLoad = _rangeMasterStep2.Range(4, 10, 205, 10);
            _cellMasterStep2DiffLoadLower = _rangeMasterStep2.Range(4, 11, 205, 11);
            _cellMasterStep2DiffLoadUpper = _rangeMasterStep2.Range(4, 12, 205, 12);

            _cellMasterStep2VarMap.Add(_cellMasterStep2CompStroke);
            _cellMasterStep2VarMap.Add(_cellMasterStep2CompLoad);
            _cellMasterStep2VarMap.Add(_cellMasterStep2CompLoadLower);
            _cellMasterStep2VarMap.Add(_cellMasterStep2CompLoadUpper);
            _cellMasterStep2VarMap.Add(_cellMasterStep2ExtnStroke);
            _cellMasterStep2VarMap.Add(_cellMasterStep2ExtnLoad);
            _cellMasterStep2VarMap.Add(_cellMasterStep2ExtnLoadLower);
            _cellMasterStep2VarMap.Add(_cellMasterStep2ExtnLoadUpper);
            _cellMasterStep2VarMap.Add(_cellMasterStep2DiffStroke);
            _cellMasterStep2VarMap.Add(_cellMasterStep2DiffLoad);
            _cellMasterStep2VarMap.Add(_cellMasterStep2DiffLoadLower);
            _cellMasterStep2VarMap.Add(_cellMasterStep2DiffLoadUpper);
        }

        protected List<IXLRange>? _cellMasterStep3VarMap = new List<IXLRange>();
        protected IXLRange? _cellMasterStep3CompStroke;
        protected IXLRange? _cellMasterStep3CompLoad;
        protected IXLRange? _cellMasterStep3CompLoadLower;
        protected IXLRange? _cellMasterStep3CompLoadUpper;
        protected IXLRange? _cellMasterStep3ExtnStroke;
        protected IXLRange? _cellMasterStep3ExtnLoad;
        protected IXLRange? _cellMasterStep3ExtnLoadLower;
        protected IXLRange? _cellMasterStep3ExtnLoadUpper;
        protected IXLRange? _cellMasterStep3DiffStroke;
        protected IXLRange? _cellMasterStep3DiffLoad;
        protected IXLRange? _cellMasterStep3DiffLoadLower;
        protected IXLRange? _cellMasterStep3DiffLoadUpper;

        void _initMasterStep3VarMap()
        {
            _cellMasterStep3CompStroke = _rangeMasterStep3.Range(4, 1, 205, 1);
            _cellMasterStep3CompLoad = _rangeMasterStep3.Range(4, 2, 205, 2);
            _cellMasterStep3CompLoadLower = _rangeMasterStep3.Range(4, 3, 205, 3);
            _cellMasterStep3CompLoadUpper = _rangeMasterStep3.Range(4, 4, 205, 4);
            _cellMasterStep3ExtnStroke = _rangeMasterStep3.Range(4, 5, 205, 5);
            _cellMasterStep3ExtnLoad = _rangeMasterStep3.Range(4, 6, 205, 6);
            _cellMasterStep3ExtnLoadLower = _rangeMasterStep3.Range(4, 7, 205, 7);
            _cellMasterStep3ExtnLoadUpper = _rangeMasterStep3.Range(4, 8, 205, 8);
            _cellMasterStep3DiffStroke = _rangeMasterStep3.Range(4, 9, 205, 9);
            _cellMasterStep3DiffLoad = _rangeMasterStep3.Range(4, 10, 205, 10);
            _cellMasterStep3DiffLoadLower = _rangeMasterStep3.Range(4, 11, 205, 11);
            _cellMasterStep3DiffLoadUpper = _rangeMasterStep3.Range(4, 12, 205, 12);

            _cellMasterStep3VarMap.Add(_cellMasterStep3CompStroke);
            _cellMasterStep3VarMap.Add(_cellMasterStep3CompLoad);
            _cellMasterStep3VarMap.Add(_cellMasterStep3CompLoadLower);
            _cellMasterStep3VarMap.Add(_cellMasterStep3CompLoadUpper);
            _cellMasterStep3VarMap.Add(_cellMasterStep3ExtnStroke);
            _cellMasterStep3VarMap.Add(_cellMasterStep3ExtnLoad);
            _cellMasterStep3VarMap.Add(_cellMasterStep3ExtnLoadLower);
            _cellMasterStep3VarMap.Add(_cellMasterStep3ExtnLoadUpper);
            _cellMasterStep3VarMap.Add(_cellMasterStep3DiffStroke);
            _cellMasterStep3VarMap.Add(_cellMasterStep3DiffLoad);
            _cellMasterStep3VarMap.Add(_cellMasterStep3DiffLoadLower);
            _cellMasterStep3VarMap.Add(_cellMasterStep3DiffLoadUpper);
        }

        //List<IXLCell> _cellMasterStep4VarMap = new List<IXLCell>();
        //List<IXLCell> _cellMasterStep5VarMap = new List<IXLCell>();

        protected List<IXLCell>? _cellRealtimeModelTableVarMap = new List<IXLCell>();
        protected IXLCell? _cellRealtimeModelName;
        protected IXLCell? _cellRealtimeDay;
        protected IXLCell? _cellRealtimeMonth;
        protected IXLCell? _cellRealtimeYear;
        protected IXLCell? _cellRealtimeHour;
        protected IXLCell? _cellRealtimeMinute;
        protected IXLCell? _cellRealtimeSecond;

        void _initRealtimeModelTableVarMap()
        {
            _cellRealtimeModelName = _rangeRealtimeModelTable.Cell(1, 2);
            _cellRealtimeDay = _rangeRealtimeModelTable.Cell(2, 2);
            _cellRealtimeMonth = _rangeRealtimeModelTable.Cell(2, 4);
            _cellRealtimeYear = _rangeRealtimeModelTable.Cell(2, 6);
            _cellRealtimeHour = _rangeRealtimeModelTable.Cell(3, 2);
            _cellRealtimeMinute = _rangeRealtimeModelTable.Cell(3, 4);
            _cellRealtimeSecond = _rangeRealtimeModelTable.Cell(3, 6);

            _cellRealtimeModelTableVarMap.Add(_cellRealtimeModelName);
            _cellRealtimeModelTableVarMap.Add(_cellRealtimeDay);
            _cellRealtimeModelTableVarMap.Add(_cellRealtimeMonth);
            _cellRealtimeModelTableVarMap.Add(_cellRealtimeYear);
            _cellRealtimeModelTableVarMap.Add(_cellRealtimeHour);
            _cellRealtimeModelTableVarMap.Add(_cellRealtimeMinute);
            _cellRealtimeModelTableVarMap.Add(_cellRealtimeSecond);
        }

        protected List<IXLCell>? _cellRealtimeStep1ParamVarMap = new List<IXLCell>();
        protected IXLCell? _cellRealtimeStep1Mode;
        protected IXLCell? _cellRealtimeStep1Stroke;
        protected IXLCell? _cellRealtimeStep1CompSpeed;
        protected IXLCell? _cellRealtimeStep1ExtnSpeed;
        protected IXLCell? _cellRealtimeStep1CycleCount;
        protected IXLCell? _cellRealtimeStep1MaxLoad;

        void _initRealtimeStep1ParamVarMap()
        {
            _cellRealtimeStep1Mode = _rangeRealtimeStep1Param.Cell(2, 4);
            _cellRealtimeStep1Stroke = _rangeRealtimeStep1Param.Cell(3, 4);
            _cellRealtimeStep1CompSpeed = _rangeRealtimeStep1Param.Cell(4, 4);
            _cellRealtimeStep1ExtnSpeed = _rangeRealtimeStep1Param.Cell(5, 4);
            _cellRealtimeStep1CycleCount = _rangeRealtimeStep1Param.Cell(6, 4);
            _cellRealtimeStep1MaxLoad = _rangeRealtimeStep1Param.Cell(7, 4);

            _cellRealtimeStep1ParamVarMap.Add(_cellRealtimeStep1Mode);
            _cellRealtimeStep1ParamVarMap.Add(_cellRealtimeStep1Stroke);
            _cellRealtimeStep1ParamVarMap.Add(_cellRealtimeStep1CompSpeed);
            _cellRealtimeStep1ParamVarMap.Add(_cellRealtimeStep1ExtnSpeed);
            _cellRealtimeStep1ParamVarMap.Add(_cellRealtimeStep1CycleCount);
            _cellRealtimeStep1ParamVarMap.Add(_cellRealtimeStep1MaxLoad);
        }

        protected List<IXLCell>? _cellRealtimeStep2345ParamVarMap = new List<IXLCell>();
        protected IXLCell? _cellRealtimeStep2Mode;
        protected IXLCell? _cellRealtimeStep2CompSpeed;
        protected IXLCell? _cellRealtimeStep2CompJudgePosMin;
        protected IXLCell? _cellRealtimeStep2CompJudgePosMax;
        protected IXLCell? _cellRealtimeStep2CompLoadRefPos;
        protected IXLCell? _cellRealtimeStep2ExtnSpeed;
        protected IXLCell? _cellRealtimeStep2ExtnJudgePosMin;
        protected IXLCell? _cellRealtimeStep2ExtnJudgePosMax;
        protected IXLCell? _cellRealtimeStep2ExtnLoadRefPos;
        protected IXLCell? _cellRealtimeStep2LoadRefTolerance;
        protected IXLCell? _cellRealtimeStep3Mode;
        protected IXLCell? _cellRealtimeStep3CompSpeed;
        protected IXLCell? _cellRealtimeStep3CompJudgePosMin;
        protected IXLCell? _cellRealtimeStep3CompJudgePosMax;
        protected IXLCell? _cellRealtimeStep3CompLoadRefPos;
        protected IXLCell? _cellRealtimeStep3ExtnSpeed;
        protected IXLCell? _cellRealtimeStep3ExtnJudgePosMin;
        protected IXLCell? _cellRealtimeStep3ExtnJudgePosMax;
        protected IXLCell? _cellRealtimeStep3ExtnLoadRefPos;
        protected IXLCell? _cellRealtimeStep3LoadRefTolerance;

        void _initRealtimeStep2345ParamVarMap()
        {
            _cellRealtimeStep2Mode = _rangeRealtimeStep2345Param.Cell(2, 5);
            _cellRealtimeStep2CompSpeed = _rangeRealtimeStep2345Param.Cell(3, 5);
            _cellRealtimeStep2CompJudgePosMin = _rangeRealtimeStep2345Param.Cell(4, 5);
            _cellRealtimeStep2CompJudgePosMax = _rangeRealtimeStep2345Param.Cell(5, 5);
            _cellRealtimeStep2CompLoadRefPos = _rangeRealtimeStep2345Param.Cell(6, 5);
            _cellRealtimeStep2ExtnSpeed = _rangeRealtimeStep2345Param.Cell(7, 5);
            _cellRealtimeStep2ExtnJudgePosMin = _rangeRealtimeStep2345Param.Cell(8, 5);
            _cellRealtimeStep2ExtnJudgePosMax = _rangeRealtimeStep2345Param.Cell(9, 5);
            _cellRealtimeStep2ExtnLoadRefPos = _rangeRealtimeStep2345Param.Cell(10, 5);
            _cellRealtimeStep2LoadRefTolerance = _rangeRealtimeStep2345Param.Cell(11, 5);
            _cellRealtimeStep3Mode = _rangeRealtimeStep2345Param.Cell(2, 6);
            _cellRealtimeStep3CompSpeed = _rangeRealtimeStep2345Param.Cell(3, 6);
            _cellRealtimeStep3CompJudgePosMin = _rangeRealtimeStep2345Param.Cell(4, 6);
            _cellRealtimeStep3CompJudgePosMax = _rangeRealtimeStep2345Param.Cell(5, 6);
            _cellRealtimeStep3CompLoadRefPos = _rangeRealtimeStep2345Param.Cell(6, 6);
            _cellRealtimeStep3ExtnSpeed = _rangeRealtimeStep2345Param.Cell(7, 6);
            _cellRealtimeStep3ExtnJudgePosMin = _rangeRealtimeStep2345Param.Cell(8, 6);
            _cellRealtimeStep3ExtnJudgePosMax = _rangeRealtimeStep2345Param.Cell(9, 6);
            _cellRealtimeStep3ExtnLoadRefPos = _rangeRealtimeStep2345Param.Cell(10, 6);
            _cellRealtimeStep3LoadRefTolerance = _rangeRealtimeStep2345Param.Cell(11, 6);

            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2Mode);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2CompSpeed);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2CompJudgePosMin);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2CompJudgePosMax);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2CompLoadRefPos);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2ExtnSpeed);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2ExtnJudgePosMin);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2ExtnJudgePosMax);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2ExtnLoadRefPos);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep2LoadRefTolerance);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3Mode);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3CompSpeed);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3CompJudgePosMin);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3CompJudgePosMax);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3CompLoadRefPos);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3ExtnSpeed);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3ExtnJudgePosMin);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3ExtnJudgePosMax);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3ExtnLoadRefPos);
            _cellRealtimeStep2345ParamVarMap.Add(_cellRealtimeStep3LoadRefTolerance);
        }

        protected List<IXLCell>? _cellRealtimeJudgementVarMap = new List<IXLCell>();
        protected IXLCell? _cellMaxLoad;
        protected IXLCell? _cellStep2CompLoadRef;
        protected IXLCell? _cellStep2ExtnLoadRef;
        protected IXLCell? _cellStep3CompLoadRef;
        protected IXLCell? _cellStep3ExtnLoadRef;

        void _initRealtimeJudgementVarMap()
        {
            _cellMaxLoad = _rangeRealtimeJudgement.Cell(3, 3);
            _cellStep2CompLoadRef = _rangeRealtimeJudgement.Cell(4, 4);
            _cellStep2ExtnLoadRef = _rangeRealtimeJudgement.Cell(5, 4);
            _cellStep3CompLoadRef = _rangeRealtimeJudgement.Cell(4, 5);
            _cellStep3ExtnLoadRef = _rangeRealtimeJudgement.Cell(5, 5);

            _cellRealtimeJudgementVarMap.Add(_cellMaxLoad);
            _cellRealtimeJudgementVarMap.Add(_cellStep2CompLoadRef);
            _cellRealtimeJudgementVarMap.Add(_cellStep2ExtnLoadRef);
            _cellRealtimeJudgementVarMap.Add(_cellStep3CompLoadRef);
            _cellRealtimeJudgementVarMap.Add(_cellStep3ExtnLoadRef);
        }

        protected List<IXLRange>? _cellRealtimeStep2VarMap = new List<IXLRange>();
        protected IXLRange? _cellRealtimeStep2CompStroke;
        protected IXLRange? _cellRealtimeStep2CompLoad;
        protected IXLRange? _cellRealtimeStep2ExtnStroke;
        protected IXLRange? _cellRealtimeStep2ExtnLoad;
        protected IXLRange? _cellRealtimeStep2DiffStroke;
        protected IXLRange? _cellRealtimeStep2DiffLoad;

        void _initRealtimeStep2VarMap()
        {
            _cellRealtimeStep2CompStroke = _rangeRealtimeStep2.Range(4, 1, 205, 1);
            _cellRealtimeStep2CompLoad = _rangeRealtimeStep2.Range(4, 2, 205, 2);
            _cellRealtimeStep2ExtnStroke = _rangeRealtimeStep2.Range(4, 3, 205, 3);
            _cellRealtimeStep2ExtnLoad = _rangeRealtimeStep2.Range(4, 4, 205, 4);
            _cellRealtimeStep2DiffStroke = _rangeRealtimeStep2.Range(4, 5, 205, 5);
            _cellRealtimeStep2DiffLoad = _rangeRealtimeStep2.Range(4, 6, 205, 6);

            _cellRealtimeStep2VarMap.Add(_cellRealtimeStep2CompStroke);
            _cellRealtimeStep2VarMap.Add(_cellRealtimeStep2CompLoad);
            _cellRealtimeStep2VarMap.Add(_cellRealtimeStep2ExtnStroke);
            _cellRealtimeStep2VarMap.Add(_cellRealtimeStep2ExtnLoad);
            _cellRealtimeStep2VarMap.Add(_cellRealtimeStep2DiffStroke);
            _cellRealtimeStep2VarMap.Add(_cellRealtimeStep2DiffLoad);
        }

        protected List<IXLRange>? _cellRealtimeStep3VarMap = new List<IXLRange>();
        protected IXLRange? _cellRealtimeStep3CompStroke;
        protected IXLRange? _cellRealtimeStep3CompLoad;
        protected IXLRange? _cellRealtimeStep3ExtnStroke;
        protected IXLRange? _cellRealtimeStep3ExtnLoad;
        protected IXLRange? _cellRealtimeStep3DiffStroke;
        protected IXLRange? _cellRealtimeStep3DiffLoad;

        void _initRealtimeStep3VarMap()
        {
            _cellRealtimeStep3CompStroke = _rangeRealtimeStep3.Range(4, 1, 205, 1);
            _cellRealtimeStep3CompLoad = _rangeRealtimeStep3.Range(4, 2, 205, 2);
            _cellRealtimeStep3ExtnStroke = _rangeRealtimeStep3.Range(4, 3, 205, 3);
            _cellRealtimeStep3ExtnLoad = _rangeRealtimeStep3.Range(4, 4, 205, 4);
            _cellRealtimeStep3DiffStroke = _rangeRealtimeStep3.Range(4, 5, 205, 5);
            _cellRealtimeStep3DiffLoad = _rangeRealtimeStep3.Range(4, 6, 205, 6);

            _cellRealtimeStep3VarMap.Add(_cellRealtimeStep3CompStroke);
            _cellRealtimeStep3VarMap.Add(_cellRealtimeStep3CompLoad);
            _cellRealtimeStep3VarMap.Add(_cellRealtimeStep3ExtnStroke);
            _cellRealtimeStep3VarMap.Add(_cellRealtimeStep3ExtnLoad);
            _cellRealtimeStep3VarMap.Add(_cellRealtimeStep3DiffStroke);
            _cellRealtimeStep3VarMap.Add(_cellRealtimeStep3DiffLoad);
        }

        //List<IXLCell> _cellRealtimeStep4VarMap = new List<IXLCell>();
        //List<IXLCell> _cellRealtimeStep5VarMap = new List<IXLCell>();     

        List<object[]> blueprintModelTable = new List<object[]>
            {
                new object[] { "Model:"},
                new object[] { "Date:", "", "Day", "", "Month", "", "Year"},
                new object[] { "Time:", "", "Hour", "", "Minutes", "", "Second" }
            };

        List<object[]> blueprintStep1Table = new List<object[]>
            {
                new object[] { "PARAMETERS", "", "", "STEP1"},
                new object[] { "Mode(0:Disable,1:Enable):", "", ""},
                new object[] { "Product Stroke mm:", "", ""},
                new object[] { "Compress Speed mm/s:", "", ""},
                new object[] { "Extension Speed mm/s:", "", ""},
                new object[] { "Ext/Comp Cycle t:", "", ""},
                new object[] { "Product Maximum Load:", "", ""}
            };

        List<object[]> blueprintStep2345Table = new List<object[]>
            {
                new object[] { "PARAMETERS", "", "", "", "STEP2", "STEP3" },
                new object[] { "Mode(0:Disable,1:Enable):", "", "", ""},
                new object[] { "Compress Speed mm/s:", "", "", "" },
                new object[] { "Compres  Judge Stroke Min mm:", "", "", ""},
                new object[] { "Compres  Judge Stroke Max mm:", "", "", ""},
                new object[] { "Compress Load Reference Stroke mm:", "", "", ""},
                new object[] { "Extension Speed mm/s:", "", "", "" },
                new object[] { "Extension  Judge Stroke Min mm:", "", "", ""},
                new object[] { "Extension  Judge Stroke Max mm:", "", "", ""},
                new object[] { "Extension Load Reference Stroke mm:", "", "", ""},
                new object[] { "Comp /Ext Load Reference Tolerance %:", "", "", ""}
            };

        List<object[]> blueprintMasterDataHeader = new List<object[]>
            {
                new object[] { "Master Compress/Extension Recipe Data" }
            };

        List<object[]> blueprintMasterStep2 = new List<object[]>
            {
                new object[] { "Step2 Master Data" },
                new object[] { "Compress" ," " ," " ," " ,"Extension" ," " ," " ," " ,"Difference" ," " ," " ," "},
                new object[] { "Stroke", "Master", "Lower", "Upper", "Stroke", "Master", "Lower", "Upper", "Stroke", "Master", "Lower", "Upper" }
            };

        List<object[]> blueprintMasterStep3 = new List<object[]>
            {
                new object[] { "Step3 Master Data" },
                new object[] { "Compress" ," " ," " ," " ,"Extension" ," " ," " ," " ,"Difference" ," " ," " ," "},
                new object[] { "Stroke", "Master", "Lower", "Upper", "Stroke", "Master", "Lower", "Upper", "Stroke", "Master", "Lower", "Upper" }
            };

        List<object[]> blueprintMasterStep4 = new List<object[]>
            {
                new object[] { "Step4 Master Data" },
                new object[] { "Compress" ," " ," " ," " ,"Extension" ," " ," " ," " ,"Difference" ," " ," " ," "},
                new object[] { "Stroke", "Master", "Lower", "Upper", "Stroke", "Master", "Lower", "Upper", "Stroke", "Master", "Lower", "Upper" }
            };

        List<object[]> blueprintMasterStep5 = new List<object[]>
            {
                new object[] { "Step5 Master Data" },
                new object[] { "Compress" ," " ," " ," " ,"Extension" ," " ," " ," " ,"Difference" ," " ," " ," "},
                new object[] { "Stroke", "Master", "Lower", "Upper", "Stroke", "Master", "Lower", "Upper", "Stroke", "Master", "Lower", "Upper" }
            };

        List<object[]> blueprintRealtimeJudgement = new List<object[]>
            {
                new object[] { "Product General Judgement" },
                new object[] { "", "", "STEP 1", "STEP 2", "STEP 3", "STEP 4", "STEP 5" },
                new object[] { "Maximum Load N:" },
                new object[] { "Comp Load Ref N:" },
                new object[] { "Ext Load Ref N:" }
            };

        List<object[]> blueprintRealtimeDataHeader = new List<object[]>
            {
                new object[] { "Product Compress/Extension Data Logging" }
            };

        List<object[]> blueprintRealtimeStep2 = new List<object[]>
            {
                new object[] { "Step2 Realtime Data" },
                new object[] { "Compress", "", "Extension", "", "Difference", "" },
                new object[] { "Stroke", "Load", "Stroke", "Load", "Stroke", "Load" }
            };

        List<object[]> blueprintRealtimeStep3 = new List<object[]>
            {
                new object[] { "Step3 Realtime Data" },
                new object[] { "Compress", "", "Extension", "", "Difference", "" },
                new object[] { "Stroke", "Load", "Stroke", "Load", "Stroke", "Load" }
            };

        List<object[]> blueprintRealtimeStep4 = new List<object[]>
            {
                new object[] { "Step4 Realtime Data" },
                new object[] { "Compress", "", "Extension", "", "Difference", "" },
                new object[] { "Stroke", "Load", "Stroke", "Load", "Stroke", "Load" }
            };

        List<object[]> blueprintRealtimeStep5 = new List<object[]>
            {
                new object[] { "Step5 Realtime Data" },
                new object[] { "Compress", "", "Extension", "", "Difference", "" },
                new object[] { "Stroke", "Load", "Stroke", "Load", "Stroke", "Load" }
            };

        void formattingModelTable(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintModelTable);
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.FirstColumn().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.FirstColumn().Style.Font.SetBold(true);

            wsr.Rows(2, 3).Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            wsr.Rows(2, 3).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            wsr.FirstColumn().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
        }

        void formattingStep1Param(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintStep1Table);
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.FirstRow().Style.Font.SetBold();
            wsr.Cell(1, 4).Style.Font.SetUnderline();
            wsr.Range(1, 1, 1, 3).Merge();
            wsr.Range(1, 1, 1, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.Range(2, 1, 2, 3).Merge();
            wsr.Range(2, 1, 2, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(3, 1, 3, 3).Merge();
            wsr.Range(3, 1, 3, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(4, 1, 4, 3).Merge();
            wsr.Range(4, 1, 4, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(5, 1, 5, 3).Merge();
            wsr.Range(5, 1, 5, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(6, 1, 6, 3).Merge();
            wsr.Range(6, 1, 6, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(7, 1, 7, 3).Merge();
            wsr.Range(7, 1, 7, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

            wsr.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            wsr.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            wsr.FirstRow().Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Column(3).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
        }

        void formattingStep2345Param(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintStep2345Table);
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.FirstRow().Style.Font.SetBold();
            wsr.Range(1, 5, 1, 6).Style.Font.SetUnderline();
            wsr.Range(1, 1, 1, 4).Merge();
            wsr.Range(1, 1, 1, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.Range(2, 1, 2, 4).Merge();
            wsr.Range(2, 1, 2, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(3, 1, 3, 4).Merge();
            wsr.Range(3, 1, 3, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(4, 1, 4, 4).Merge();
            wsr.Range(4, 1, 4, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(5, 1, 5, 4).Merge();
            wsr.Range(5, 1, 5, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(6, 1, 6, 4).Merge();
            wsr.Range(6, 1, 6, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(7, 1, 7, 4).Merge();
            wsr.Range(7, 1, 7, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(8, 1, 8, 4).Merge();
            wsr.Range(8, 1, 8, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(9, 1, 9, 4).Merge();
            wsr.Range(9, 1, 9, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(10, 1, 10, 4).Merge();
            wsr.Range(10, 1, 10, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(11, 1, 11, 4).Merge();
            wsr.Range(11, 1, 11, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

            wsr.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            wsr.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            wsr.FirstRow().Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Column(4).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
        }

        void formattingMasterDataHeader(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintMasterDataHeader);
            wsr.Merge();
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.Style.Font.SetBold();
            wsr.Style.Font.SetUnderline();
        }

        void formattingMasterStep2(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintMasterStep2);
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.FirstRow().Style.Font.SetBold();
            wsr.FirstRow().Merge();
            wsr.Range(2, 1, 2, 4).Merge();
            wsr.Range(2, 5, 2, 8).Merge();
            wsr.Range(2, 9, 2, 12).Merge();

            wsr.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            wsr.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            wsr.FirstRow().Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Row(2).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Row(3).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Column(4).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
            wsr.Column(8).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
        }

        void formattingMasterStep3(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintMasterStep3);
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.FirstRow().Style.Font.SetBold();
            wsr.FirstRow().Merge();
            wsr.Range(2, 1, 2, 4).Merge();
            wsr.Range(2, 5, 2, 8).Merge();
            wsr.Range(2, 9, 2, 12).Merge();

            wsr.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            wsr.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            wsr.FirstRow().Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Row(2).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Row(3).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Column(4).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
            wsr.Column(8).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
        }

        //void formattingMasterStep4(IXLRange wsr) { }
        //void formattingMasterStep5(IXLRange wsr) { }

        void formattingRealtimeJudgement(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintRealtimeJudgement);
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.FirstRow().Style.Font.SetBold();
            wsr.FirstRow().Merge();
            wsr.Row(2).Style.Font.SetBold();
            wsr.Range(2, 3, 2, 7).Style.Font.SetUnderline();
            wsr.Range(2, 1, 2, 2).Merge();
            wsr.Range(3, 1, 3, 2).Merge();
            wsr.Range(3, 1, 3, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(4, 1, 4, 2).Merge();
            wsr.Range(4, 1, 4, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            wsr.Range(5, 1, 5, 2).Merge();
            wsr.Range(5, 1, 5, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

            wsr.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            wsr.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            wsr.FirstRow().Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Row(2).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Column(2).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
        }

        void formattingRealtimeDataHeader(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintRealtimeDataHeader);
            wsr.Merge();
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.Style.Font.SetBold();
            wsr.Style.Font.SetUnderline();
        }

        void formattingRealtimeStep2(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintRealtimeStep2);
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.FirstRow().Style.Font.SetBold();
            wsr.FirstRow().Merge();
            wsr.Range(2, 1, 2, 2).Merge();
            wsr.Range(2, 3, 2, 4).Merge();
            wsr.Range(2, 5, 2, 6).Merge();

            wsr.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            wsr.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            wsr.FirstRow().Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Row(2).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Row(3).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Column(2).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
            wsr.Column(4).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
        }

        void formattingRealtimeStep3(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintRealtimeStep3);
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.FirstRow().Style.Font.SetBold();
            wsr.FirstRow().Merge();
            wsr.Range(2, 1, 2, 2).Merge();
            wsr.Range(2, 3, 2, 4).Merge();
            wsr.Range(2, 5, 2, 6).Merge();

            wsr.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            wsr.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            wsr.FirstRow().Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Row(2).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Row(3).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
            wsr.Column(2).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
            wsr.Column(4).Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
        }

        //void formattingRealtimeStep4(IXLRange wsr) { }
        //void formattingRealtimeStep5(IXLRange wsr) { }

        public XLBLUEPRINT()
        {
            _mastering = _XLblueprint.AddWorksheet("Master Blueprint");
            _rangeMasterModelTable = _mastering.Range("A1:G3");
            formattingModelTable(ref _rangeMasterModelTable);
            _initMasterModelTableVarMap();
            _rangeMasterStep1Param = _mastering.Range("A6:D12");
            formattingStep1Param(ref _rangeMasterStep1Param);
            _initMasterStep1ParamVarMap();
            _rangeMasterStep2345Param = _mastering.Range("F6:M16");
            formattingStep2345Param(ref _rangeMasterStep2345Param);
            _initMasterStep2345ParamVarMap();
            _rangeMasterDataHeader = _mastering.Range("A18:Y18");
            formattingMasterDataHeader(ref _rangeMasterDataHeader);
            _rangeMasterStep2 = _mastering.Range("A19:L223");
            formattingMasterStep2(ref _rangeMasterStep2);
            _initMasterStep2VarMap();
            _rangeMasterStep3 = _mastering.Range("N19:Y223");
            formattingMasterStep3(ref _rangeMasterStep3);
            _initMasterStep3VarMap();
            //_rangeMasterStep4
            //_rangeMasterStep5

            _realtime = _XLblueprint.AddWorksheet("Realtime Blueprint");
            _rangeRealtimeModelTable = _realtime.Range("A1:G3");
            formattingModelTable(ref _rangeRealtimeModelTable);
            _initRealtimeModelTableVarMap();
            _rangeRealtimeStep1Param = _realtime.Range("A6:D12");
            formattingStep1Param(ref _rangeRealtimeStep1Param);
            _initRealtimeStep1ParamVarMap();
            _rangeRealtimeStep2345Param = _realtime.Range("F6:M16");
            formattingStep2345Param(ref _rangeRealtimeStep2345Param);
            _initRealtimeStep2345ParamVarMap();
            _rangeRealtimeJudgement = _realtime.Range("A18:G22");
            formattingRealtimeJudgement(ref _rangeRealtimeJudgement);
            _initRealtimeJudgementVarMap();
            _rangeRealtimeDataHeader = _realtime.Range("A24:M24");
            formattingRealtimeDataHeader(ref _rangeRealtimeDataHeader);
            _rangeRealtimeStep2 = _realtime.Range("A25:F229");
            formattingRealtimeStep2(ref _rangeRealtimeStep2);
            _initRealtimeStep2VarMap();
            _rangeRealtimeStep3 = _realtime.Range("H25:M229");
            formattingRealtimeStep3(ref _rangeRealtimeStep3);
            _initRealtimeStep3VarMap();
            //_rangeRealtimeStep4
            //_rangeRealtimeStep5
        }

        public int TemplatePrint(string filename)
        {
            try { _XLblueprint.SaveAs(filename); }
            catch { }
            finally { }
            return 1;
        }

        public XLWorkbook GetTemplateWB() { return _XLblueprint; }
        public IXLWorksheet GetMasterWS() { return _mastering; }
        public IXLWorksheet GetRealtimeWS() { return _realtime; }

    }

    public class EXCELSTREAM : XLBLUEPRINT
    {
        int? _filemode;
        public EXCELSTREAM(string WBTYPE)
        {
            if (WBTYPE == "MASTER")
            {
                _XLblueprint.Worksheets.Delete(2);
                _filemode = 0;
            }
            else if (WBTYPE == "REALTIME")
            {
                _XLblueprint.Worksheets.Delete(1);
                _filemode = 1;
            }
        }

        public int FilePrint(string filename)
        {
            try { base._XLblueprint.SaveAs(filename); return 1; }
            catch { return 0; }
        }

        public int FileRead(ref XLWorkbook wbObject, string filename)
        {
            try { wbObject = new XLWorkbook(filename); return 1; }
            catch { return 0; }
        }


        public int setModelName(string modelname)
        {
            //try
            {
                if (_filemode == 1) { string sbuff = new string(modelname); _cellRealtimeModelTableVarMap[0].SetValue(sbuff); }
                else if (_filemode == 0) { string sbuff = new string(modelname); _cellMasterModelTableVarMap[0].SetValue(sbuff); }
                return 1;
            }
            //catch{ return 0; }
        }

        public string getModelName()
        {
            try
            {
                if (_filemode == 1) { return _cellRealtimeModelTableVarMap[0].GetString(); }
                else if (_filemode == 0) { return _cellMasterModelTableVarMap[0].GetString(); }
                else { return " "; }
            }
            catch { return " "; }
        }

        public int setDateTime(List<string> buffer)
        {
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < 6; i++) { string sbuff = new string(buffer[i]); _cellRealtimeModelTableVarMap[i+1].SetValue(sbuff); }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < 6; i++) { string sbuff = new string(buffer[i]); _cellMasterModelTableVarMap[i + 1].SetValue(sbuff); }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<string> getDateTime()
        {
            List<string> buffer = new List<string> { };
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < 6; i++) { buffer.Add(_cellRealtimeModelTableVarMap[i + 1].GetString()); }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < 6; i++) { buffer.Add(_cellMasterModelTableVarMap[i + 1].GetString()); }
                }
            }
            catch {  }
            return buffer;
        }

        public int setParameterStep1(List<string> buffer)
        {
            //try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeStep1ParamVarMap.Count; i++) { string sbuff = new string(buffer[i]); _cellRealtimeStep1ParamVarMap[i].SetValue(sbuff); }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < _cellMasterStep1ParamVarMap.Count; i++) { string sbuff = new string(buffer[i]); _cellMasterStep1ParamVarMap[i].SetValue(sbuff); }
                }
                return 1;
            }
            //catch { return 0; }

        }

        public List<string> getParameterStep1()
        {
            List<string> buffer = new List<string> { };
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeStep1ParamVarMap.Count; i++) { buffer.Add(_cellRealtimeStep1ParamVarMap[i].GetString()); }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < _cellMasterStep1ParamVarMap.Count; i++) { buffer.Add(_cellMasterStep1ParamVarMap[i].GetString()); }
                }
            }
            catch { }
            return buffer;
        }

        public int setParameterStep2345(List<string> buffer)
        {
            //try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeStep2345ParamVarMap.Count; i++) { string sbuff = new string(buffer[i]); _cellRealtimeStep2345ParamVarMap[i].SetValue(sbuff); }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < _cellMasterStep2345ParamVarMap.Count; i++) { string sbuff = new string(buffer[i]); _cellMasterStep2345ParamVarMap[i].SetValue(sbuff); }
                }
                return 1;
            }
            //catch { return 0; }
        }

        public List<string> getParameterStep2345()
        {
            List<string> buffer = new List<string> { };
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeStep2345ParamVarMap.Count; i++) { buffer.Add(_cellRealtimeStep2345ParamVarMap[i].GetString()); }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < _cellMasterStep2345ParamVarMap.Count; i++) { buffer.Add(_cellMasterStep2345ParamVarMap[i].GetString()); }
                }
            }
            catch { }
            return buffer;
        }

        public int setMasterStep2(List<List<string>> buffer)
        {
            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < _cellMasterStep2VarMap.Count; iv++)
                    {
                        List<string> scope = buffer[iv];
                        for (int ivy = 0; ivy < _cellMasterStep2VarMap[iv].RowCount(); ivy++)
                        {
                            _cellMasterStep2VarMap[iv].Row(ivy).SetValue(scope[ivy]);
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<List<string>> getMasterStep2()
        {
            List<List<string>> buffer = new List<List<string>> { };
            List<string> buffercompressstroke = new List<string> { };
            List<string> buffercompressload = new List<string> { };
            List<string> buffercompresslower = new List<string> { };
            List<string> buffercompressupper = new List<string> { };
            List<string> bufferextendstroke = new List<string> { };
            List<string> bufferextendload = new List<string> { };
            List<string> bufferextendlower = new List<string> { };
            List<string> bufferextendupper = new List<string> { };
            List<string> bufferdifferencestroke = new List<string> { };
            List<string> bufferdifferenceload = new List<string> { };
            List<string> bufferdifferencelower = new List<string> { };
            List<string> bufferdifferenceupper = new List<string> { };

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < _cellMasterStep2VarMap.Count; iv++)
                    {
                        List<string> scope = new List<string>();
                        for (int ivy = 0; ivy < _cellMasterStep2VarMap[iv].RowCount(); ivy++)
                        {
                            scope.Add(_cellMasterStep2VarMap[iv].Row(ivy).ToString());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }

        public int setMasterStep3(List<List<string>> buffer)
        {
            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < _cellMasterStep3VarMap.Count; iv++)
                    {
                        List<string> scope = buffer[iv];
                        for (int ivy = 0; ivy < _cellMasterStep3VarMap[iv].RowCount(); ivy++)
                        {
                            _cellMasterStep3VarMap[iv].Row(ivy).SetValue(scope[ivy]);
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<List<string>> getMasterStep3()
        {
            List<List<string>> buffer = new List<List<string>> { };
            List<string> buffercompressstroke = new List<string> { };
            List<string> buffercompressload = new List<string> { };
            List<string> buffercompresslower = new List<string> { };
            List<string> buffercompressupper = new List<string> { };
            List<string> bufferextendstroke = new List<string> { };
            List<string> bufferextendload = new List<string> { };
            List<string> bufferextendlower = new List<string> { };
            List<string> bufferextendupper = new List<string> { };
            List<string> bufferdifferencestroke = new List<string> { };
            List<string> bufferdifferenceload = new List<string> { };
            List<string> bufferdifferencelower = new List<string> { };
            List<string> bufferdifferenceupper = new List<string> { };

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < _cellMasterStep3VarMap.Count; iv++)
                    {
                        List<string> scope = new List<string>();
                        for (int ivy = 0; ivy < _cellMasterStep3VarMap[iv].RowCount(); ivy++)
                        {
                            scope.Add(_cellMasterStep3VarMap[iv].Row(ivy).ToString());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }

        public int setRealtimeJudgement(List<string> buffer)
        {
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeJudgementVarMap.Count; i++) { _cellRealtimeJudgementVarMap[i].SetValue(buffer[i]); }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<string> getRealtimeJudgement()
        {
            List<string> buffer = new List<string> { };
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeJudgementVarMap.Count; i++) { buffer.Add(_cellRealtimeJudgementVarMap[i].GetString()); }
                }
            }
            catch { }
            return buffer;
        }

        public int setRealtimeStep2(List<List<string>> buffer)
        {
            try
            {
                if (_filemode == 1)
                {
                    for (int iv = 0; iv < _cellRealtimeStep2VarMap.Count; iv++)
                    {
                        List<string> scope = buffer[iv];
                        for (int ivy = 0; ivy < _cellRealtimeStep2VarMap[iv].RowCount(); ivy++)
                        {
                            _cellRealtimeStep2VarMap[iv].Row(ivy).SetValue(scope[ivy]);
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<List<string>> getRealtimeStep2()
        {
            List<List<string>> buffer = new List<List<string>> { };
            List<string> buffercompressstroke = new List<string> { };
            buffer.Add(buffercompressstroke);
            List<string> buffercompressload = new List<string> { };
            buffer.Add(buffercompressload);
            List<string> bufferextendstroke = new List<string> { };
            buffer.Add(bufferextendstroke);
            List<string> bufferextendload = new List<string> { };
            buffer.Add(bufferextendload);
            List<string> bufferdifferencestroke = new List<string> { };
            buffer.Add(bufferdifferencestroke);
            List<string> bufferdifferenceload = new List<string> { };
            buffer.Add(bufferdifferenceload);

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < _cellRealtimeStep2VarMap.Count; iv++)
                    {
                        List<string> scope = new List<string>();
                        for (int ivy = 0; ivy < _cellRealtimeStep2VarMap[iv].RowCount(); ivy++)
                        {
                            scope.Add(_cellRealtimeStep2VarMap[iv].Row(ivy).ToString());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }

        public int setRealtimeStep3(List<List<string>> buffer)
        {
            try
            {
                if (_filemode == 1)
                {
                    for (int iv = 0; iv < _cellRealtimeStep3VarMap.Count; iv++)
                    {
                        List<string> scope = buffer[iv];
                        for (int ivy = 0; ivy < _cellRealtimeStep3VarMap[iv].RowCount(); ivy++)
                        {
                            _cellRealtimeStep3VarMap[iv].Row(ivy).SetValue(scope[ivy]);
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<List<string>> getRealtimeStep3()
        {
            List<List<string>> buffer = new List<List<string>> { };
            List<string> buffercompressstroke = new List<string> { };
            buffer.Add(buffercompressstroke);
            List<string> buffercompressload = new List<string> { };
            buffer.Add(buffercompressload);
            List<string> bufferextendstroke = new List<string> { };
            buffer.Add(bufferextendstroke);
            List<string> bufferextendload = new List<string> { };
            buffer.Add(bufferextendload);
            List<string> bufferdifferencestroke = new List<string> { };
            buffer.Add(bufferdifferencestroke);
            List<string> bufferdifferenceload = new List<string> { };
            buffer.Add(bufferdifferenceload);

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < _cellRealtimeStep3VarMap.Count; iv++)
                    {
                        List<string> scope = new List<string>();
                        for (int ivy = 0; ivy < _cellRealtimeStep3VarMap[iv].RowCount(); ivy++)
                        {
                            scope.Add(_cellRealtimeStep3VarMap[iv].Row(ivy).ToString());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            EXCELSTREAM MasterFile1 = new EXCELSTREAM("MASTER");
            EXCELSTREAM RealtimeFile1 = new EXCELSTREAM("REALTIME");
        }
    }
}
