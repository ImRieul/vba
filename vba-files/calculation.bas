Attribute VB_Name = "calculation"

public sub divideRange(ByVal rng as Range, ByVal Optional div as Long)
    dim cell as Range
    dim i as Integer
    
    if div = 0 then
        div = 1
    end if
    
    for each cell in rng
        if Not tools.isFormulaInCell(cell) then

            if (cell.value = 0 or not IsNumeric(cell.value)) then
                cell.value = 0
            else
                cell.value = cell.value / div
            end if

        end if
    next cell
end sub
