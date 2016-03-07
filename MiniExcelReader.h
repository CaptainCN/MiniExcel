#ifndef _TINYXLSX_H_
#define _TINYXLSX_H_

#include <vector>

namespace MiniExcelReader
{
    struct Cell
    {
        std::string value;
        std::string type;
    };

    struct Range
    {
        int firstRow;
        int lastRow;
        int firstCol;
        int lastCol;
    };

    struct Sheet
    {
        bool visible;

        Range dimension;

        std::string name;
        std::string path;
        
        Cell* getCell(int row, int col);

    private:
        friend class ExcelFile;

        int sheetId;
        std::string rid;
        std::vector<Cell*> _cells;
    };

    class Zip;

    class ExcelFile
    {
    public:
        ~ExcelFile();
        bool open(const char* filename);

        Sheet* getSheet(const char* name);
        std::vector<Sheet>& sheets() { return _sheets; }

    private:

        void readWorkBook(const char* filename);
        void readWorkBookRels(const char* filename);
        void readSharedStrings(const char* filename);
        void readStyles(const char* filename);
        void readSheet(Sheet& sh);

        void parseCell(const std::string& value, int& row, int& col);
        void parseRange(const std::string& value, Range& range);

        std::vector<std::string> _sharedString;
        std::vector<Sheet> _sheets;
        Zip* _zip;
    };
}


#endif