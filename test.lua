local miniexcel = require "miniexcel"

local x = miniexcel.open('1.xlsx')

local sheets = x:sheets()

print(type(sheets), #sheets)

for i, sheet1 in ipairs(sheets) do

    local f = io.open(sheet1:name() .. '.txt', 'w')
    print(sheet1:name())

    local dim = sheet1:dimension();

    for r = dim.firstRow, dim.lastRow do
        local rowTable = {}
        for c = dim.firstCol, dim.lastCol do
                local cell = sheet1:cell(r, c)

                local str = "."
                if cell then
                    str = cell.value
                end

                table.insert(rowTable, string.format("%30s", str))
        end
        f:write(table.concat(rowTable, '|') .. '\n')
    end

    f:close()
end

