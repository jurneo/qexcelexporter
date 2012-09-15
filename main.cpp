// ExcelModule.cpp : Defines the entry point for the console application.
//
#include "stdafx.h"
#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <cmath>

#include "excelmodule.h"
#include <basicexcel.h>

#include <model/peakoutput.h>

#include <boost/format.hpp>
#include <boost/algorithm/string.hpp>
#include <boost/lexical_cast.hpp>
#include <boost/shared_ptr.hpp>


typedef boost::shared_ptr<PeakOutput> PeakOutputPtr;

namespace
{
const int START_ROW = 0;
const char* SHEET_NAME = "TIR_Results";

std::string cell(int row, int col)
{
    if (row < 0 || col < 0)
    {
        return "a1";
    }

    std::string value = boost::str(boost::format("%s%d") % (char)(97 + col) % (row + 1));
    return value;
}

std::string narrow_string1(const std::wstring& str)
{
    std::string ret;

    if (!str.empty())
    {
        ret.resize(str.length());

        typedef std::ctype<wchar_t> CT;
        CT const& ct = std::_USE(std::locale(), CT);

        ct.narrow(&str[0], &*str.begin() + str.size(), '?', &ret[0]);
    }

    return ret;
}

std::wstring widen_string(const std::string& str)
{
    std::wstring ret;

    if (!str.empty())
    {
        ret.resize(str.length());

        typedef std::ctype<wchar_t> CT;
        CT const& ct = std::_USE(std::locale(), CT);

        ct.widen(&str[0], &*str.begin() + str.size(), &ret[0]);
    }

    return ret;
}
}

void displayHelp();

bool readFromFile(const _TCHAR* fileName, std::vector<PeakOutput>& pks, std::vector<std::vector<PeakOutputPtr> >& peaks)
{
    ifstream fin(fileName, std::ios_base::in);
    if (fin.fail())  return false;

    // read the number of peak
    std::string strNumPeak;
    std::string strNumRawData;
    fin >> strNumPeak;
    fin >> strNumRawData;

    int numPeak = boost::lexical_cast<int>(strNumPeak);
    int numRawData = boost::lexical_cast<int>(strNumRawData);

    if (numPeak <= 0 || numPeak >= 1e4 || numRawData <= 0 || numRawData >= 1e4)
    {
        return false;
    }
    
    std::string m_peak;
    std::string m_peakToPeak;
    std::string m_effectivePeak;
    std::string m_averagePeak;
    std::string m_minPeak;

    for(int i = 0; i<numPeak; ++i)
    {
        PeakOutput data;
        fin >> m_peak >> m_peakToPeak >> m_effectivePeak >> m_averagePeak >> m_minPeak;
        try
        {
            data.m_peak = boost::lexical_cast<double>(m_peak);
            data.m_peakToPeak = boost::lexical_cast<double>(m_peakToPeak);
            data.m_effectivePeak = boost::lexical_cast<double>(m_effectivePeak);
            data.m_averagePeak = boost::lexical_cast<double>(m_averagePeak);
            data.m_minPeak = boost::lexical_cast<double>(m_minPeak);
        }
        catch(boost::bad_lexical_cast& e)
        {
            return false;
        }
        catch(...)
        {
            return false;
        }
        pks.push_back(data);
    }

    for(int i = 0; i<numPeak; ++i)
    {
        std::vector<PeakOutputPtr> pk2;
        for(int j = 0; j<numRawData; ++j)
        {
            PeakOutputPtr data = boost::make_shared<PeakOutput>();
            fin >> m_peak >> m_peakToPeak >> m_effectivePeak >> m_averagePeak >> m_minPeak;
            try
            {
                data->m_peak = boost::lexical_cast<double>(m_peak);
                data->m_peakToPeak = boost::lexical_cast<double>(m_peakToPeak);
                data->m_effectivePeak = boost::lexical_cast<double>(m_effectivePeak);
                data->m_averagePeak = boost::lexical_cast<double>(m_averagePeak);
                data->m_minPeak = boost::lexical_cast<double>(m_minPeak);
            }
            catch(boost::bad_lexical_cast& e)
            {
                return false;
            }
            catch(...)
            {
                return false;
            }
            pk2.push_back(data);
        }
        peaks.push_back(pk2);
    }

    return true;
}

int _tmain(int argc, _TCHAR* argv[])
{
    if (argc < 5)
    {        
        return 0;
    }
    
    bool toAppend = (std::wstring(argv[3]) == L"true");
    bool isExist = (std::wstring(argv[4]) == L"true");

    std::wstring tmpFile3(argv[1]);
    std::string tmpFile(narrow_string1(tmpFile3));

    std::vector<PeakOutput> pks;
    std::vector<std::vector<PeakOutputPtr> > peaks;
    bool status = readFromFile(argv[2], pks, peaks);

    if (!status)
    {
        return 0;
    }

    ExcelModule excel;
    int sheetIndex = 1;
    int startRow = 0;

    if (toAppend)
    {
        bool hasSheet = false;
        {
            YExcel::BasicExcel workbook;
            workbook.Load(tmpFile.c_str());
            YExcel::BasicExcelWorksheet* ws = workbook.GetWorksheet(SHEET_NAME);
            excel.getExcelSheetIndex(SHEET_NAME, sheetIndex);

            if (sheetIndex >= 1)
            {
                hasSheet = true;
            }
            else
            {
                sheetIndex = 1;
            }

            if (ws != 0) // find the starting row to append the new content
            {
                for (size_t row = 0; row < 5e4; ++row)
                {
                    int v = ws->Cell(START_ROW + 1 + row, 0)->GetInteger();

                    if (v < 1)
                    {
                        break;
                    }

                    startRow = v;
                }
            }
        }
        excel.openExcelBook(tmpFile);

        if (!hasSheet)
        {
            excel.setExcelSheetName(sheetIndex, SHEET_NAME);
            excel.setExcelValue(cell(START_ROW, 1), "Avg.TIR", false, 1);
            excel.setExcelValue(cell(START_ROW, 2), "Min.TIR", false, 1);
            excel.setExcelValue(cell(START_ROW, 3), "Max.TIR", false, 1);
            excel.setExcelValue(cell(START_ROW, 4), "RPM", false, 1);
        }
    }
    else
    {
        if (!isExist)
        {
            excel.openExcelBook(tmpFile);
        }
        else
        {
            int totalrows = -1, totalcolumns = -1;
            {
                YExcel::BasicExcel workbook;
                workbook.Load(tmpFile.c_str());
                YExcel::BasicExcelWorksheet* ws = workbook.GetWorksheet(SHEET_NAME);

                if (ws != 0)
                {
                    totalrows = ws->GetTotalRows();
                    totalcolumns = max(15, ws->GetTotalCols());
                }
            }
            excel.openExcelBook(tmpFile);

            if (totalcolumns > 0 && totalrows > 0)
            {
                for (int row = 0; row < totalrows; ++row)
                {
                    for (int col = 0; col < totalcolumns; ++col)
                    {
                        excel.setExcelValue(cell(row, col), "", false, 1);
                    }
                }
            }
        }

        excel.setExcelSheetName(sheetIndex, SHEET_NAME);
        excel.setExcelValue(cell(START_ROW, 1), "Avg.TIR", false, 1);
        excel.setExcelValue(cell(START_ROW, 2), "Min.TIR", false, 1);
        excel.setExcelValue(cell(START_ROW, 3), "Max.TIR", false, 1);
        excel.setExcelValue(cell(START_ROW, 4), "RPM", false, 1);
    }

    if (peaks.size() != pks.size())
    {
        return 0;
    }

    int mxcell = 0;

    for (size_t row = 0; row < pks.size(); ++row)
    {
        int rows = startRow + START_ROW + 1 + row;
        //ws->Cell(rows,0)->SetDouble(rows);
        excel.setExcelValue(sheetIndex, cell(rows, 0), boost::str(boost::format("%d") % rows), true, 1);

        // TODO this is a hack on pk side -- should use a separate result model..
        //ws->Cell(rows, 1)->SetDouble(pks[row].m_peakToPeak);
        excel.setExcelValue(sheetIndex, cell(rows, 1), boost::str(boost::format("%.4f") % pks[row].m_peakToPeak), true, 1);

        //ws->Cell(rows, 2)->SetDouble(pks[row].m_averagePeak);
        excel.setExcelValue(sheetIndex, cell(rows, 2), boost::str(boost::format("%.4f") % pks[row].m_averagePeak), true, 1);

        //ws->Cell(rows, 3)->SetDouble(pks[row].m_effectivePeak);
        excel.setExcelValue(sheetIndex, cell(rows, 3), boost::str(boost::format("%.4f") % pks[row].m_effectivePeak), true, 1);

        //ws->Cell(rows, 4)->SetDouble(pks[row].m_peak);
        excel.setExcelValue(sheetIndex, cell(rows, 4), boost::str(boost::format("%.4f") % pks[row].m_peak), true, 1);

        std::vector<PeakOutputPtr> raw = peaks.at(row);
        size_t rs = raw.size();

        for (size_t column = 0; column < rs; ++column)
        {
            if (column < rs)
            {
                //ws->Cell(rows, 5+column)->SetDouble(raw[column]->m_peakToPeak);
                excel.setExcelValue(sheetIndex, cell(rows, 5 + column),
                                    boost::str(boost::format("%.4f") % raw[column]->m_peakToPeak), true, 1);
            }
            else
            {
                //ws->Cell(rows, 5+column)->EraseContents();  // clear the contents.
                excel.setExcelValue(sheetIndex, cell(rows, 5 + column), "", false, 1);
            }

            if (mxcell < column)
            {
                mxcell = column;
            }
        }
    }

    for (int j = 0; j < mxcell + 1; ++j)
    {
        //ws->Cell(START_ROW, 5+j)->SetString("TIR");
        excel.setExcelValue(sheetIndex, cell(START_ROW, 5 + j), "TIR", false, 1);
    }

    if (isExist)
    {
        excel.save();
    }
    else
    {
        excel.saveAs(tmpFile.c_str());
    }

    DeleteFile(argv[2]);

    return 0;
}

