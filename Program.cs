using ClosedXML.Excel;
using ClosedXML.Report.Utils;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace LIBEXCELMANIPULATOR
{
    public class XLBLUEPRINT
    {
        protected XLWorkbook? _XLblueprint = new XLWorkbook();
        protected IXLWorksheet? _mastering;
        protected IXLWorksheet? _realtime;
        protected IXLWorksheet? _realtimeLogBuffer;

        protected IXLRange? _rangeMasterModelTable;
        protected IXLRange? _rangeMasterStep1Param;
        protected IXLRange? _rangeMasterStep2345Param;
        protected IXLRange? _rangeRsideMasterDataHeader;
        protected IXLRange? _rangeLsideMasterDataHeader;
        protected IXLRange? _rangeRsideMasterStep2;
        protected IXLRange? _rangeRsideMasterStep3;
        protected IXLRange? _rangeLsideMasterStep2;
        protected IXLRange? _rangeLsideMasterStep3;
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

        protected IXLRange? _rangeNGLABEL;
        protected IXLRange? _cellNGLABEL;

        protected IXLRange? _rangeNGLABELLogBuffer;
        protected IXLRange? _cellNGLABELLogBuffer;

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
            _cellMasterModelTableVarMap.Add(_cellMasterYear);
            _cellMasterModelTableVarMap.Add(_cellMasterMonth);
            _cellMasterModelTableVarMap.Add(_cellMasterDay);
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

        protected List<IXLRange>? _cellRsideMasterStep2VarMap = new List<IXLRange>();
        protected IXLRange? _cellRsideMasterStep2CompStroke;
        protected IXLRange? _cellRsideMasterStep2CompLoad;
        protected IXLRange? _cellRsideMasterStep2CompLoadLower;
        protected IXLRange? _cellRsideMasterStep2CompLoadUpper;
        protected IXLRange? _cellRsideMasterStep2ExtnStroke;
        protected IXLRange? _cellRsideMasterStep2ExtnLoad;
        protected IXLRange? _cellRsideMasterStep2ExtnLoadLower;
        protected IXLRange? _cellRsideMasterStep2ExtnLoadUpper;
        protected IXLRange? _cellRsideMasterStep2DiffStroke;
        protected IXLRange? _cellRsideMasterStep2DiffLoad;
        protected IXLRange? _cellRsideMasterStep2DiffLoadLower;
        protected IXLRange? _cellRsideMasterStep2DiffLoadUpper;

        protected List<IXLRange>? _cellLsideMasterStep2VarMap = new List<IXLRange>();
        protected IXLRange? _cellLsideMasterStep2CompStroke;
        protected IXLRange? _cellLsideMasterStep2CompLoad;
        protected IXLRange? _cellLsideMasterStep2CompLoadLower;
        protected IXLRange? _cellLsideMasterStep2CompLoadUpper;
        protected IXLRange? _cellLsideMasterStep2ExtnStroke;
        protected IXLRange? _cellLsideMasterStep2ExtnLoad;
        protected IXLRange? _cellLsideMasterStep2ExtnLoadLower;
        protected IXLRange? _cellLsideMasterStep2ExtnLoadUpper;
        protected IXLRange? _cellLsideMasterStep2DiffStroke;
        protected IXLRange? _cellLsideMasterStep2DiffLoad;
        protected IXLRange? _cellLsideMasterStep2DiffLoadLower;
        protected IXLRange? _cellLsideMasterStep2DiffLoadUpper;

        void _initMasterStep2VarMap()
        {
            _cellRsideMasterStep2CompStroke = _rangeRsideMasterStep2.Range(4, 1, 205, 1);
            _cellRsideMasterStep2CompLoad = _rangeRsideMasterStep2.Range(4, 2, 205, 2);
            _cellRsideMasterStep2CompLoadLower = _rangeRsideMasterStep2.Range(4, 3, 205, 3);
            _cellRsideMasterStep2CompLoadUpper = _rangeRsideMasterStep2.Range(4, 4, 205, 4);
            _cellRsideMasterStep2ExtnStroke = _rangeRsideMasterStep2.Range(4, 5, 205, 5);
            _cellRsideMasterStep2ExtnLoad = _rangeRsideMasterStep2.Range(4, 6, 205, 6);
            _cellRsideMasterStep2ExtnLoadLower = _rangeRsideMasterStep2.Range(4, 7, 205, 7);
            _cellRsideMasterStep2ExtnLoadUpper = _rangeRsideMasterStep2.Range(4, 8, 205, 8);
            _cellRsideMasterStep2DiffStroke = _rangeRsideMasterStep2.Range(4, 9, 205, 9);
            _cellRsideMasterStep2DiffLoad = _rangeRsideMasterStep2.Range(4, 10, 205, 10);
            _cellRsideMasterStep2DiffLoadLower = _rangeRsideMasterStep2.Range(4, 11, 205, 11);
            _cellRsideMasterStep2DiffLoadUpper = _rangeRsideMasterStep2.Range(4, 12, 205, 12);

            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2CompStroke);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2CompLoad);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2CompLoadLower);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2CompLoadUpper);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2ExtnStroke);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2ExtnLoad);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2ExtnLoadLower);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2ExtnLoadUpper);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2DiffStroke);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2DiffLoad);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2DiffLoadLower);
            _cellRsideMasterStep2VarMap.Add(_cellRsideMasterStep2DiffLoadUpper);

            _cellLsideMasterStep2CompStroke = _rangeLsideMasterStep2.Range(4, 1, 205, 1);
            _cellLsideMasterStep2CompLoad = _rangeLsideMasterStep2.Range(4, 2, 205, 2);
            _cellLsideMasterStep2CompLoadLower = _rangeLsideMasterStep2.Range(4, 3, 205, 3);
            _cellLsideMasterStep2CompLoadUpper = _rangeLsideMasterStep2.Range(4, 4, 205, 4);
            _cellLsideMasterStep2ExtnStroke = _rangeLsideMasterStep2.Range(4, 5, 205, 5);
            _cellLsideMasterStep2ExtnLoad = _rangeLsideMasterStep2.Range(4, 6, 205, 6);
            _cellLsideMasterStep2ExtnLoadLower = _rangeLsideMasterStep2.Range(4, 7, 205, 7);
            _cellLsideMasterStep2ExtnLoadUpper = _rangeLsideMasterStep2.Range(4, 8, 205, 8);
            _cellLsideMasterStep2DiffStroke = _rangeLsideMasterStep2.Range(4, 9, 205, 9);
            _cellLsideMasterStep2DiffLoad = _rangeLsideMasterStep2.Range(4, 10, 205, 10);
            _cellLsideMasterStep2DiffLoadLower = _rangeLsideMasterStep2.Range(4, 11, 205, 11);
            _cellLsideMasterStep2DiffLoadUpper = _rangeLsideMasterStep2.Range(4, 12, 205, 12);

            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2CompStroke);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2CompLoad);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2CompLoadLower);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2CompLoadUpper);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2ExtnStroke);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2ExtnLoad);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2ExtnLoadLower);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2ExtnLoadUpper);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2DiffStroke);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2DiffLoad);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2DiffLoadLower);
            _cellLsideMasterStep2VarMap.Add(_cellLsideMasterStep2DiffLoadUpper);
        }

        protected List<IXLRange>? _cellRsideMasterStep3VarMap = new List<IXLRange>();
        protected IXLRange? _cellRsideMasterStep3CompStroke;
        protected IXLRange? _cellRsideMasterStep3CompLoad;
        protected IXLRange? _cellRsideMasterStep3CompLoadLower;
        protected IXLRange? _cellRsideMasterStep3CompLoadUpper;
        protected IXLRange? _cellRsideMasterStep3ExtnStroke;
        protected IXLRange? _cellRsideMasterStep3ExtnLoad;
        protected IXLRange? _cellRsideMasterStep3ExtnLoadLower;
        protected IXLRange? _cellRsideMasterStep3ExtnLoadUpper;
        protected IXLRange? _cellRsideMasterStep3DiffStroke;
        protected IXLRange? _cellRsideMasterStep3DiffLoad;
        protected IXLRange? _cellRsideMasterStep3DiffLoadLower;
        protected IXLRange? _cellRsideMasterStep3DiffLoadUpper;

        protected List<IXLRange>? _cellLsideMasterStep3VarMap = new List<IXLRange>();
        protected IXLRange? _cellLsideMasterStep3CompStroke;
        protected IXLRange? _cellLsideMasterStep3CompLoad;
        protected IXLRange? _cellLsideMasterStep3CompLoadLower;
        protected IXLRange? _cellLsideMasterStep3CompLoadUpper;
        protected IXLRange? _cellLsideMasterStep3ExtnStroke;
        protected IXLRange? _cellLsideMasterStep3ExtnLoad;
        protected IXLRange? _cellLsideMasterStep3ExtnLoadLower;
        protected IXLRange? _cellLsideMasterStep3ExtnLoadUpper;
        protected IXLRange? _cellLsideMasterStep3DiffStroke;
        protected IXLRange? _cellLsideMasterStep3DiffLoad;
        protected IXLRange? _cellLsideMasterStep3DiffLoadLower;
        protected IXLRange? _cellLsideMasterStep3DiffLoadUpper;

        void _initMasterStep3VarMap()
        {
            _cellRsideMasterStep3CompStroke = _rangeRsideMasterStep3.Range(4, 1, 205, 1);
            _cellRsideMasterStep3CompLoad = _rangeRsideMasterStep3.Range(4, 2, 205, 2);
            _cellRsideMasterStep3CompLoadLower = _rangeRsideMasterStep3.Range(4, 3, 205, 3);
            _cellRsideMasterStep3CompLoadUpper = _rangeRsideMasterStep3.Range(4, 4, 205, 4);
            _cellRsideMasterStep3ExtnStroke = _rangeRsideMasterStep3.Range(4, 5, 205, 5);
            _cellRsideMasterStep3ExtnLoad = _rangeRsideMasterStep3.Range(4, 6, 205, 6);
            _cellRsideMasterStep3ExtnLoadLower = _rangeRsideMasterStep3.Range(4, 7, 205, 7);
            _cellRsideMasterStep3ExtnLoadUpper = _rangeRsideMasterStep3.Range(4, 8, 205, 8);
            _cellRsideMasterStep3DiffStroke = _rangeRsideMasterStep3.Range(4, 9, 205, 9);
            _cellRsideMasterStep3DiffLoad = _rangeRsideMasterStep3.Range(4, 10, 205, 10);
            _cellRsideMasterStep3DiffLoadLower = _rangeRsideMasterStep3.Range(4, 11, 205, 11);
            _cellRsideMasterStep3DiffLoadUpper = _rangeRsideMasterStep3.Range(4, 12, 205, 12);

            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3CompStroke);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3CompLoad);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3CompLoadLower);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3CompLoadUpper);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3ExtnStroke);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3ExtnLoad);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3ExtnLoadLower);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3ExtnLoadUpper);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3DiffStroke);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3DiffLoad);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3DiffLoadLower);
            _cellRsideMasterStep3VarMap.Add(_cellRsideMasterStep3DiffLoadUpper);

            _cellLsideMasterStep3CompStroke = _rangeLsideMasterStep3.Range(4, 1, 205, 1);
            _cellLsideMasterStep3CompLoad = _rangeLsideMasterStep3.Range(4, 2, 205, 2);
            _cellLsideMasterStep3CompLoadLower = _rangeLsideMasterStep3.Range(4, 3, 205, 3);
            _cellLsideMasterStep3CompLoadUpper = _rangeLsideMasterStep3.Range(4, 4, 205, 4);
            _cellLsideMasterStep3ExtnStroke = _rangeLsideMasterStep3.Range(4, 5, 205, 5);
            _cellLsideMasterStep3ExtnLoad = _rangeLsideMasterStep3.Range(4, 6, 205, 6);
            _cellLsideMasterStep3ExtnLoadLower = _rangeLsideMasterStep3.Range(4, 7, 205, 7);
            _cellLsideMasterStep3ExtnLoadUpper = _rangeLsideMasterStep3.Range(4, 8, 205, 8);
            _cellLsideMasterStep3DiffStroke = _rangeLsideMasterStep3.Range(4, 9, 205, 9);
            _cellLsideMasterStep3DiffLoad = _rangeLsideMasterStep3.Range(4, 10, 205, 10);
            _cellLsideMasterStep3DiffLoadLower = _rangeLsideMasterStep3.Range(4, 11, 205, 11);
            _cellLsideMasterStep3DiffLoadUpper = _rangeLsideMasterStep3.Range(4, 12, 205, 12);

            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3CompStroke);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3CompLoad);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3CompLoadLower);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3CompLoadUpper);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3ExtnStroke);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3ExtnLoad);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3ExtnLoadLower);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3ExtnLoadUpper);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3DiffStroke);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3DiffLoad);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3DiffLoadLower);
            _cellLsideMasterStep3VarMap.Add(_cellLsideMasterStep3DiffLoadUpper);
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
            _cellRealtimeModelTableVarMap.Add(_cellRealtimeYear);
            _cellRealtimeModelTableVarMap.Add(_cellRealtimeMonth);
            _cellRealtimeModelTableVarMap.Add(_cellRealtimeDay);
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

        List<object[]> blueprintRsideMasterDataHeader = new List<object[]>
            {
                new object[] { "R side Master Compress/Extension Recipe Data" }
            };

        List<object[]> blueprintLsideMasterDataHeader = new List<object[]>
            {
                new object[] { "L side Master Compress/Extension Recipe Data" }
            };

        List<object[]> blueprintMasterStep2 = new List<object[]>
            {
                new object[] { "Step2 Master Data" },
                new object[] { "Compress" ," " ," " ," " ,"Extension" ," " ," " ," " ,"Difference" ," " ," " ," "},
                new object[] { "Stroke", "Load", "Lower", "Upper", "Stroke", "Load", "Lower", "Upper", "Stroke", "Load", "Lower", "Upper" }
            };

        List<object[]> blueprintMasterStep3 = new List<object[]>
            {
                new object[] { "Step3 Master Data" },
                new object[] { "Compress" ," " ," " ," " ,"Extension" ," " ," " ," " ,"Difference" ," " ," " ," "},
                new object[] { "Stroke", "Load", "Lower", "Upper", "Stroke", "Load", "Lower", "Upper", "Stroke", "Load", "Lower", "Upper" }
            };

        List<object[]> blueprintMasterStep4 = new List<object[]>
            {
                new object[] { "Step4 Master Data" },
                new object[] { "Compress" ," " ," " ," " ,"Extension" ," " ," " ," " ,"Difference" ," " ," " ," "},
                new object[] { "Stroke", "Load", "Lower", "Upper", "Stroke", "Load", "Lower", "Upper", "Stroke", "Load", "Lower", "Upper" }
            };

        List<object[]> blueprintMasterStep5 = new List<object[]>
            {
                new object[] { "Step5 Master Data" },
                new object[] { "Compress" ," " ," " ," " ,"Extension" ," " ," " ," " ,"Difference" ," " ," " ," "},
                new object[] { "Stroke", "Load", "Lower", "Upper", "Stroke", "Load", "Lower", "Upper", "Stroke", "Load", "Lower", "Upper" }
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

        void formattingRsideMasterDataHeader(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintRsideMasterDataHeader);
            wsr.Merge();
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.Style.Font.SetBold();
            wsr.Style.Font.SetUnderline();
        }

        void formattingLsideMasterDataHeader(ref IXLRange wsr)
        {
            wsr.FirstCell().InsertData(blueprintLsideMasterDataHeader);
            wsr.Merge();
            wsr.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Bottom);
            wsr.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsr.Style.Font.SetBold();
            wsr.Style.Font.SetUnderline();
        }

        void formattingRsideMasterStep2(ref IXLRange wsr)
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

        void formattingRsideMasterStep3(ref IXLRange wsr)
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

        void formattingLsideMasterStep2(ref IXLRange wsr)
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

        void formattingLsideMasterStep3(ref IXLRange wsr)
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
            //Master Data
            _mastering = _XLblueprint.AddWorksheet("Master DATA");
            //Common
            _rangeMasterModelTable = _mastering.Range("A1:G3");
            formattingModelTable(ref _rangeMasterModelTable);
            _initMasterModelTableVarMap();
            _rangeMasterStep1Param = _mastering.Range("A6:D12");
            formattingStep1Param(ref _rangeMasterStep1Param);
            _initMasterStep1ParamVarMap();
            _rangeMasterStep2345Param = _mastering.Range("F6:M16");
            formattingStep2345Param(ref _rangeMasterStep2345Param);
            _initMasterStep2345ParamVarMap();
            //Header R
            _rangeRsideMasterDataHeader = _mastering.Range("A18:Y18");
            formattingRsideMasterDataHeader(ref _rangeRsideMasterDataHeader);
            //Header L
            _rangeLsideMasterDataHeader = _mastering.Range("AA18:AY18");
            formattingLsideMasterDataHeader(ref _rangeLsideMasterDataHeader);
            //Step2 R
            _rangeRsideMasterStep2 = _mastering.Range("A19:L223");
            formattingRsideMasterStep2(ref _rangeRsideMasterStep2);
            //Step2 L
            _rangeLsideMasterStep2 = _mastering.Range("AA19:AL223");
            formattingLsideMasterStep2(ref _rangeLsideMasterStep2);
            //Step2 Init
            _initMasterStep2VarMap();

            //Step3 R
            _rangeRsideMasterStep3 = _mastering.Range("N19:Y223");
            formattingRsideMasterStep3(ref _rangeRsideMasterStep3);
            //Step3 L
            _rangeLsideMasterStep3 = _mastering.Range("AN19:AY223");
            formattingLsideMasterStep3(ref _rangeLsideMasterStep3);
            //Step3 Init
            _initMasterStep3VarMap();
            //_rangeMasterStep4
            //_rangeMasterStep5

            _realtime = _XLblueprint.AddWorksheet("Realtime DATA");
            _rangeNGLABEL = _realtime.Range("I2:J3");
            _rangeNGLABEL.Merge();
            _cellNGLABEL = _rangeNGLABEL;
            
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

            /*
            _realtimeLogBuffer = _XLblueprint.AddWorksheet("Realtime DATA");
            _rangeNGLABELLogBuffer = _realtimeLogBuffer.Range("I2:J3");
            _rangeNGLABELLogBuffer.Merge();
            _cellNGLABELLogBuffer = _rangeNGLABELLogBuffer;

            _rangeRealtimeModelTable = _realtimeLogBuffer.Range("A1:G3");
            formattingModelTable(ref _rangeRealtimeModelTable);
            _initRealtimeModelTableVarMap();
            _rangeRealtimeStep1Param = _realtimeLogBuffer.Range("A6:D12");
            formattingStep1Param(ref _rangeRealtimeStep1Param);
            _initRealtimeStep1ParamVarMap();
            _rangeRealtimeStep2345Param = _realtimeLogBuffer.Range("F6:M16");
            formattingStep2345Param(ref _rangeRealtimeStep2345Param);
            _initRealtimeStep2345ParamVarMap();
            _rangeRealtimeJudgement = _realtimeLogBuffer.Range("A18:G22");
            formattingRealtimeJudgement(ref _rangeRealtimeJudgement);
            _initRealtimeJudgementVarMap();
            _rangeRealtimeDataHeader = _realtimeLogBuffer.Range("A24:M24");
            formattingRealtimeDataHeader(ref _rangeRealtimeDataHeader);
            _rangeRealtimeStep2 = _realtimeLogBuffer.Range("A25:F229");
            formattingRealtimeStep2(ref _rangeRealtimeStep2);
            _initRealtimeStep2VarMap();
            _rangeRealtimeStep3 = _realtime.Range("H25:M229");
            formattingRealtimeStep3(ref _rangeRealtimeStep3);
            _initRealtimeStep3VarMap();
            //_rangeRealtimeStep4
            //_rangeRealtimeStep5
            */
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

        public int FileReadMaster(string filename)
        {
            XLWorkbook wbObject;
            IXLWorksheet wsObject;
            IXLRange rangeObject;
            IXLRange rangeReadMaster;

            try
            {
                using (FileStream stream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    wbObject = new XLWorkbook(filename);
                    wsObject = wbObject.Worksheet("Master DATA");
                    rangeObject = wsObject.Range("A1", "AZ225");

                    rangeReadMaster = _mastering.Range("A1", "AZ225");

                    foreach (var row in rangeObject.Rows())
                    {
                        foreach (var cell in row.Cells())
                        {
                            rangeReadMaster.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber).Value = cell.Value;
                        }
                    }
                }

                return 1;
            }
            catch
            {
                return 0;
            }
        }

        public int FileReadRealtime(string filename)
        {
            XLWorkbook wbObject;
            IXLWorksheet wsObject;
            IXLRange rangeObject;
            IXLRange rangeReadRealtime;

            try
            {
                using (FileStream stream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    wbObject = new XLWorkbook(filename);
                    wsObject = wbObject.Worksheet("Master DATA");
                    rangeObject = wsObject.Range("A1", "Z225");

                    rangeReadRealtime = _realtime.Range("A1", "Z225");

                    foreach (var row in rangeObject.Rows())
                    {
                        foreach (var cell in row.Cells())
                        {
                            rangeReadRealtime.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber).Value = cell.Value;
                        }
                    }
                }

                return 1;
            }
            catch
            {
                return 0;
            }
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

        public int setDateTime(List<String> buffer)
        {
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < 6; i++) { var sbuff = buffer[i].ToString(); _cellRealtimeModelTableVarMap[i+1].SetValue((XLCellValue)sbuff); }
                }
                
                else if (_filemode == 0)
                {
                    for (int i = 0; i < 6; i++) { var sbuff = buffer[i].ToString(); _cellMasterModelTableVarMap[i + 1].SetValue((XLCellValue)sbuff); }
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

        public int setParameterStep1doublediscrete(double buffer, int index)
        {
            //try
            {
                if (_filemode == 1)
                {
                    _cellRealtimeStep1ParamVarMap[index].SetValue(buffer);
}
                else if (_filemode == 0)
                {
                    _cellMasterStep1ParamVarMap[index].SetValue(buffer);
                }
                return 1;
            }
            //catch { return 0; }

        }

        public double getParameterStep1doublediscrete(int index)
        {
            double buffer = new double();
            byte buff = new byte();
            bool b = new bool(); 
            try
            {
                if (_filemode == 1)
                {
                    b = _cellRealtimeStep1ParamVarMap[index].TryGetValue<double>(out buffer);
                }
                else if (_filemode == 0)
                {
                    b = _cellMasterStep1ParamVarMap[index].TryGetValue<double>(out buffer);
                }
            }
            catch { }
            return buffer;
        }

        public int setParameterStep1<T>(List<T> buffer)
        {
            //try
            {
                List<Object?> buffobj = buffer.ConvertAll(x => (Object)x);

                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeStep1ParamVarMap.Count; i++) 
                    { 
                        var sbuff = Convert.ChangeType(buffobj[i], buffobj[i].GetType());
                        Int32 check_int32 = new Int32();
                        Int16 check_int16 = new Int16();
                        Single check_float = new Single();

                        if (sbuff.GetType() != check_float.GetType())
                        {
                            _cellRealtimeStep1ParamVarMap[i].SetValue((Single)Convert.ToSingle(sbuff));
                        }
                        else if (sbuff.GetType() == check_float.GetType())
                        {
                            _cellRealtimeStep1ParamVarMap[i].SetValue((Single)sbuff);
                        }
                    }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < _cellMasterStep1ParamVarMap.Count; i++)
                    {
                        var sbuff = Convert.ChangeType(buffobj[i], buffobj[i].GetType());
                        Int32 check_int32 = new Int32();
                        Int16 check_int16 = new Int16();
                        Single check_float = new Single();

                        if (sbuff.GetType() != check_float.GetType())
                        {
                            _cellMasterStep1ParamVarMap[i].SetValue((Single)Convert.ToSingle(sbuff));
                        }
                        else if (sbuff.GetType() == check_float.GetType())
                        {
                            _cellMasterStep1ParamVarMap[i].SetValue((Single)sbuff);
                        }
                    }
                }
                return 1;
            }
            //catch { return 0; }

        }

        public List<object> getParameterStep1()
        {
            List<object> buffer = new List<object> { };
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeStep1ParamVarMap.Count; i++) { buffer.Add(_cellRealtimeStep1ParamVarMap[i]); }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < _cellMasterStep1ParamVarMap.Count; i++) { buffer.Add(_cellMasterStep1ParamVarMap[i]); }
                }
            }
            catch { }
            return buffer;
        }

        public int setParameterStep2345<T>(List<T> buffer)
        {
            //try
            List<Object?> buffobj = buffer.ConvertAll(x => (Object)x);
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeStep2345ParamVarMap.Count; i++) 
                    { 
                        var sbuff = Convert.ChangeType(buffobj[i], buffobj[i].GetType());
                        Int32 check_int = new Int32();
                        Single check_float = new Single();
                        if (sbuff.GetType() != check_float.GetType())
                        {
                            _cellRealtimeStep2345ParamVarMap[i].SetValue((Single)Convert.ToSingle(sbuff));
                        }
                        else if (sbuff.GetType() == check_float.GetType())
                        {
                            _cellRealtimeStep2345ParamVarMap[i].SetValue((Single)sbuff);
                        }
                    }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < _cellMasterStep2345ParamVarMap.Count; i++)
                    {
                        var sbuff = Convert.ChangeType(buffobj[i], buffobj[i].GetType());
                        Int32 check_int = new Int32();
                        Single check_float = new Single();
                        if (sbuff.GetType() != check_float.GetType())
                        {
                            _cellMasterStep2345ParamVarMap[i].SetValue((Single)Convert.ToSingle(sbuff));
                        }
                        else if (sbuff.GetType() == check_float.GetType())
                        {
                            _cellMasterStep2345ParamVarMap[i].SetValue((Single)sbuff);
                        }
                    }
                }
                return 1;
            }
            //catch { return 0; }
        }

        public List<object> getParameterStep2345()
        {
            List<object> buffer = new List<object> { };
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeStep2345ParamVarMap.Count; i++) { buffer.Add(_cellRealtimeStep2345ParamVarMap[i]); }
                }
                else if (_filemode == 0)
                {
                    for (int i = 0; i < _cellMasterStep2345ParamVarMap.Count; i++) { buffer.Add(_cellMasterStep2345ParamVarMap[i]); }
                }
            }
            catch { }
            return buffer;
        }

        public int setRsideMasterStep2<T>(List<List<T>> buffer)
        {
            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellRsideMasterStep2VarMap.Count); iv++)
                    {
                        List<Object?> scope = buffer[iv].ConvertAll(x => (Object)x);
                        for (int ivy = 0; ivy < (scope.Count - 1); ivy++)
                        {
                            _cellRsideMasterStep2VarMap[iv].Row(ivy + 1).SetValue((Single)Convert.ChangeType(scope[ivy], scope[ivy].GetType()));
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public int setLsideMasterStep2<T>(List<List<T>> buffer)
        {
            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellLsideMasterStep2VarMap.Count); iv++)
                    {
                        List<Object?> scope = buffer[iv].ConvertAll(x => (Object)x);
                        for (int ivy = 0; ivy < (scope.Count - 1); ivy++)
                        {
                            _cellLsideMasterStep2VarMap[iv].Row(ivy + 1).SetValue((Single)Convert.ChangeType(scope[ivy], scope[ivy].GetType()));
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<List<object>> getRsideMasterStep2()
        {
            List<List<object>> buffer = new List<List<object>> { };

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellRsideMasterStep2VarMap.Count - 1); iv++)
                    {
                        List<object> scope = new List<object>();
                        for (int ivy = 0; ivy < (_cellRsideMasterStep2VarMap[iv].RowCount() - 1); ivy++)
                        {
                            scope.Add(_cellRsideMasterStep2VarMap[iv].Row(ivy + 1).As<Object>());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }

        public List<List<object>> getLsideMasterStep2()
        {
            List<List<object>> buffer = new List<List<object>> { };

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellLsideMasterStep2VarMap.Count - 1); iv++)
                    {
                        List<object> scope = new List<object>();
                        for (int ivy = 0; ivy < (_cellLsideMasterStep2VarMap[iv].RowCount() - 1); ivy++)
                        {
                            scope.Add(_cellLsideMasterStep2VarMap[iv].Row(ivy + 1).As<Object>());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }

        public int setRsideMasterStep3<T>(List<List<T>> buffer)
        {
            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellRsideMasterStep3VarMap.Count - 1); iv++)
                    {
                        List<Object?> scope = buffer[iv].ConvertAll(x => (Object)x);
                        for (int ivy = 0; ivy < (scope.Count - 1); ivy++)
                        {
                            _cellRsideMasterStep3VarMap[iv].Row(ivy + 1).SetValue((Single)Convert.ChangeType(scope[ivy], scope[ivy].GetType()));
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public int setLsideMasterStep3<T>(List<List<T>> buffer)
        {
            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellLsideMasterStep3VarMap.Count - 1); iv++)
                    {
                        List<Object?> scope = buffer[iv].ConvertAll(x => (Object)x);
                        for (int ivy = 0; ivy < (scope.Count - 1); ivy++)
                        {
                            _cellLsideMasterStep3VarMap[iv].Row(ivy + 1).SetValue((Single)Convert.ChangeType(scope[ivy], scope[ivy].GetType()));
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<List<object>> getRsideMasterStep3()
        {
            List<List<object>> buffer = new List<List<object>> { };

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellRsideMasterStep3VarMap.Count - 1); iv++)
                    {
                        List<object> scope = new List<object>();
                        for (int ivy = 0; ivy < (_cellRsideMasterStep3VarMap[iv].RowCount() - 1); ivy++)
                        {
                            scope.Add(_cellRsideMasterStep3VarMap[iv].Row(ivy + 1).As<Object>());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }

        public List<List<object>> getLsideMasterStep3()
        {
            List<List<object>> buffer = new List<List<object>> { };

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellLsideMasterStep3VarMap.Count - 1); iv++)
                    {
                        List<object> scope = new List<object>();
                        for (int ivy = 0; ivy < (_cellLsideMasterStep3VarMap[iv].RowCount() - 1); ivy++)
                        {
                            scope.Add(_cellLsideMasterStep3VarMap[iv].Row(ivy + 1).As<Object>());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }


        public int setRealtimeJudgement<Single>(List<Single> buffer)
        {
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeJudgementVarMap.Count; i++) { _cellRealtimeJudgementVarMap[i].SetValue((float)Convert.ToSingle(buffer[i])); }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<object> getRealtimeJudgement()
        {
            List<object> buffer = new List<object> { };
            try
            {
                if (_filemode == 1)
                {
                    for (int i = 0; i < _cellRealtimeJudgementVarMap.Count; i++) { buffer.Add(_cellRealtimeJudgementVarMap[i].As<Object>()); }
                }
            }
            catch { }
            return buffer;
        }

        public int setRealtimeStep2<T>(List<List<T>> buffer)
        {
            try
            {
                if (_filemode == 1)
                {
                    for (int iv = 0; iv < (_cellRealtimeStep2VarMap.Count); iv++)
                    {
                        List<Object?> scope = buffer[iv].ConvertAll(x => (Object)x);
                        for (int ivy = 0; ivy < (scope.Count - 1); ivy++)
                        {
                            _cellRealtimeStep2VarMap[iv].Row(ivy + 1).SetValue((Single)Convert.ChangeType(scope[ivy], scope[ivy].GetType()));
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<List<object>> getRealtimeStep2()
        {
            List<List<object>> buffer = new List<List<object>> { };
            List<object> buffercompressstroke = new List<object> { };
            buffer.Add(buffercompressstroke);
            List<object> buffercompressload = new List<object> { };
            buffer.Add(buffercompressload);
            List<object> bufferextendstroke = new List<object> { };
            buffer.Add(bufferextendstroke);
            List<object> bufferextendload = new List<object> { };
            buffer.Add(bufferextendload);
            List<object> bufferdifferencestroke = new List<object> { };
            buffer.Add(bufferdifferencestroke);
            List<object> bufferdifferenceload = new List<object> { };
            buffer.Add(bufferdifferenceload);

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellRealtimeStep2VarMap.Count); iv++)
                    {
                        List<object> scope = new List<object>();
                        for (int ivy = 0; ivy < (_cellRealtimeStep2VarMap[iv].RowCount() - 1); ivy++)
                        {
                            scope.Add(_cellRealtimeStep2VarMap[iv].Row(ivy + 1).As<Object>());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }

        public int setRealtimeStep3<T>(List<List<T>> buffer)
        {
            try
            {
                if (_filemode == 1)
                {
                    for (int iv = 0; iv < (_cellRealtimeStep3VarMap.Count); iv++)
                    {
                        List<Object?> scope = buffer[iv].ConvertAll(x => (Object)x);
                        for (int ivy = 0; ivy < (scope.Count - 1); ivy++)
                        {
                            _cellRealtimeStep3VarMap[iv].Row(ivy + 1).SetValue((Single)Convert.ChangeType(scope[ivy], scope[ivy].GetType()));
                        }
                    }
                }
                return 1;
            }
            catch { return 0; }
        }

        public List<List<object>> getRealtimeStep3()
        {
            List<List<object>> buffer = new List<List<object>> { };
            List<object> buffercompressstroke = new List<object> { };
            buffer.Add(buffercompressstroke);
            List<object> buffercompressload = new List<object> { };
            buffer.Add(buffercompressload);
            List<object> bufferextendstroke = new List<object> { };
            buffer.Add(bufferextendstroke);
            List<object> bufferextendload = new List<object> { };
            buffer.Add(bufferextendload);
            List<object> bufferdifferencestroke = new List<object> { };
            buffer.Add(bufferdifferencestroke);
            List<object> bufferdifferenceload = new List<object> { };
            buffer.Add(bufferdifferenceload);

            try
            {
                if (_filemode == 0)
                {
                    for (int iv = 0; iv < (_cellRealtimeStep3VarMap.Count); iv++)
                    {
                        List<object> scope = new List<object>();
                        for (int ivy = 0; ivy < (_cellRealtimeStep3VarMap[iv].RowCount() - 1); ivy++)
                        {
                            scope.Add(_cellRealtimeStep3VarMap[iv].Row(ivy + 1).As<Object>());
                        }
                        buffer.Add(scope);
                    }
                }
            }
            catch { }
            return buffer;
        }

        public void RESET_LABEL_NG()
        {
            _cellNGLABEL.SetValue("");
            _cellNGLABEL.Style.Fill.SetBackgroundColor(XLColor.Transparent);
            _cellMaxLoad.Style.Border.SetOutsideBorderColor(XLColor.Black);
            _cellStep2CompLoadRef.Style.Border.SetOutsideBorderColor(XLColor.Black);
            _cellRealtimeStep2CompStroke.Style.Border.SetOutsideBorderColor(XLColor.Black);
            _cellRealtimeStep2CompLoad.Style.Border.SetOutsideBorderColor(XLColor.Black);
            _cellStep2ExtnLoadRef.Style.Border.SetOutsideBorderColor(XLColor.Black);
            _cellRealtimeStep2ExtnStroke.Style.Border.SetOutsideBorderColor(XLColor.Black);
            _cellRealtimeStep2ExtnLoad.Style.Border.SetOutsideBorderColor(XLColor.Black);
            _cellRealtimeStep2DiffStroke.Style.Border.SetOutsideBorderColor(XLColor.Black);
            _cellRealtimeStep2DiffLoad.Style.Border.SetOutsideBorderColor(XLColor.Black);
        }

        public void SET_LABEL_NG()
        {
            _cellNGLABEL.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            _cellNGLABEL.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            _cellNGLABEL.Style.Fill.SetBackgroundColor(XLColor.Red);
            _cellNGLABEL.Style.Font.SetBold(true);
            _cellNGLABEL.Style.Font.SetFontSize(24);
            _cellNGLABEL.SetValue("NG");
        }

    public void STEP1_MAXLOAD_NG_SET()
        {
            _cellMaxLoad.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _cellMaxLoad.Style.Border.SetOutsideBorderColor(XLColor.Red);
        }

        public void STEP2_COMP_REF_NG_SET()
        {
            _cellStep2CompLoadRef.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _cellStep2CompLoadRef.Style.Border.SetOutsideBorderColor(XLColor.Red);
        }

        public void STEP2_COMP_GRAPH_NG_SET()
        {
            _cellRealtimeStep2CompStroke.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _cellRealtimeStep2CompStroke.Style.Border.SetOutsideBorderColor(XLColor.Red);
            _cellRealtimeStep2CompLoad.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _cellRealtimeStep2CompLoad.Style.Border.SetOutsideBorderColor(XLColor.Red);
        }

        public void STEP2_EXTN_REF_NG_SET()
        {
            _cellStep2ExtnLoadRef.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _cellStep2ExtnLoadRef.Style.Border.SetOutsideBorderColor(XLColor.Red);
        }

        public void STEP2_EXTN_GRAPH_NG_SET()
        {
            _cellRealtimeStep2ExtnStroke.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _cellRealtimeStep2ExtnStroke.Style.Border.SetOutsideBorderColor(XLColor.Red);
            _cellRealtimeStep2ExtnLoad.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _cellRealtimeStep2ExtnLoad.Style.Border.SetOutsideBorderColor(XLColor.Red);
        }

        public void STEP2_DIFF_GRAPH_NG_SET()
        {
            _cellRealtimeStep2DiffStroke.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _cellRealtimeStep2DiffStroke.Style.Border.SetOutsideBorderColor(XLColor.Red);
            _cellRealtimeStep2DiffLoad.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _cellRealtimeStep2DiffLoad.Style.Border.SetOutsideBorderColor(XLColor.Red);
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            //EXCELSTREAM MasterFile1 = new EXCELSTREAM("MASTER");
            //EXCELSTREAM RealtimeFile1 = new EXCELSTREAM("REALTIME");

            //MasterFile1.setModelName("KAYABA1");
            //MasterFile1.setParameterStep1doublediscrete(990.568, 1);
            //MasterFile1.SET_LABEL_NG();

            /*
            XLWorkbook wbObjectTest;
            IXLWorksheet wsObjectTest;
            IXLRange rangeObjectTest;
            IXLRange rangeObjectModelNameTest;
            IXLCell cellObjectStrokeTest;

            Double datatest1;
            String datatest1a;
            String datatest2;

            XLWorkbook wbCopyTest;
            IXLWorksheet wsCopyTest;
            IXLRange rangeCopyTest;
            IXLRange rangeCopyModelNameTest;
            IXLRange rangeCopyStrokeTest;
            IXLCell cellCopyStrokeTest;
            */


            /*
            var excelApp = new Excel.Application();
            Excel.Workbook interopWorkbook = null;

            wbObjectTest = new XLWorkbook();
            wsObjectTest = wbObjectTest.AddWorksheet("Master DATA");

            // Attempt to get the workbook by name
            interopWorkbook = excelApp.Workbooks["TestMaster1.xlsx"];

            // Access the first worksheet
            Excel.Worksheet interopWorksheet = (Excel.Worksheet)interopWorkbook.Worksheets[1];

            // Read data from the Interop worksheet and write to ClosedXML worksheet
            Excel.Range usedRange = interopWorksheet.UsedRange;
            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    // Get the value from the Interop worksheet
                    var value = usedRange.Cells[row, col].Value;

                    // Set the value in the ClosedXML worksheet
                    wsObjectTest.Cell(row, col).Value = value;
                }
            }
            */

            //List<object> parseObjectTest = new List<object>();

            /*

            FileStream stream = new FileStream("TestMaster1.xlsx", FileMode.Open, FileAccess.Read, FileShare.Read);

            wbObjectTest = new XLWorkbook(stream);
            wsObjectTest = wbObjectTest.Worksheet("Master DATA");
            rangeObjectTest = wsObjectTest.Range("A1", "Z225");

            //rangeObjectModelNameTest = wsObjectTest.Range("A1:G3");
            //cellObjectStrokeTest = wsObjectTest.Cell("D8");

            wbCopyTest = new XLWorkbook();
            wsCopyTest = wbCopyTest.AddWorksheet("Master DATA");
            rangeCopyTest = wsCopyTest.Range("A1", "Z225");

            rangeCopyModelNameTest = wsCopyTest.Range("A1:G3");
            rangeCopyStrokeTest = wsCopyTest.Range("D8:D8");
            cellCopyStrokeTest = wsCopyTest.Range("D8:D8").Cell(1,1);

            int startRow = 1;

            foreach (var row in rangeObjectTest.RowsUsed())
            {
                // Loop through each cell in the row
                foreach (var cell in row.Cells())
                {
                    rangeCopyTest.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber).Value = cell.Value;
                }
            }


            //rangeObjectModelNameTest.FirstCell().TryGetValue<string>(out datatest2);
            //cellObjectStrokeTest.TryGetValue<double>(out datatest1);

            rangeCopyModelNameTest.Cell(1, 2).TryGetValue<string>(out datatest2);
            cellCopyStrokeTest.TryGetValue<double>(out datatest1);
            
            Console.WriteLine(datatest2);
            Console.WriteLine(datatest1);
            */

            /*
            MasterFile1.FileReadMaster("TestMaster1.xlsx");
            Console.WriteLine(MasterFile1.getModelName());
            Console.WriteLine(MasterFile1.getParameterStep1doublediscrete(1));


            RealtimeFile1.setModelName("KAYABA2");
            RealtimeFile1.setParameterStep1doublediscrete(677.568, 1);
            RealtimeFile1.RESET_LABEL_NG();
            RealtimeFile1.FilePrint("TestRealtime1.xlsx");
            Console.WriteLine(RealtimeFile1.getModelName());
            Console.WriteLine(RealtimeFile1.getParameterStep1doublediscrete(1));
            */

            //Console.ReadKey();
        }
    }
}
