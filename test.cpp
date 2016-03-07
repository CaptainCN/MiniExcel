#include "MiniExcelReader.h"

int main(int argc, const char *argv[])
{
    MiniExcelReader::ExcelFile x;
    if (!x.open("../1.xlsx"))
    {
        printf("can't open.");
        return 1;
    }

    MiniExcelReader::Sheet* sh = x.getSheet("Sheet1");

    FILE* file = fopen("../out.txt", "w");

    const MiniExcelReader::Range& dim = sh->getDimension();

    for (int r = dim.firstRow; r <= dim.lastRow; r++)
    {
        for (int c = dim.firstCol; c <= dim.lastCol; c++)
        {
            MiniExcelReader::Cell* cell = sh->getCell(r, c);

            const char* str = cell ? cell->value.c_str() : ".";

            fwrite(str, strlen(str), 1, file);
            fwrite("|", 1, 1, file);
        }
        fwrite("\n", 1, 1, file);
    }

    fclose(file);
    return 0;
}