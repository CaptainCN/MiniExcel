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

    for (int r = sh->dimension.firstRow; r <= sh->dimension.lastRow; r++)
    {
        for (int c = sh->dimension.firstCol; c <= sh->dimension.lastCol; c++)
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