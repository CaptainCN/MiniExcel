local miniexcel = require "miniexcel"

local x = miniexcel.open('1.xlsx')

local sheet1 = x:getSheet('Sheet1')


local f = io.open('outlua.txt', 'w')

local dim = sheet1:dimension();

for r = dim.firstRow, dim.lastRow do
    for c = dim.firstCol, dim.lastCol do
            local cell = sheet1:cell(r, c)

            local str = "."
            if cell then
                str = cell.value
            end

            f:write(str)
            f:write("|")
    end
    f:write('\n')
end
