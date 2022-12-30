Attribute VB_Name = "data_renewal"


public sub dataRenewalYearMonth() 

    dim data, renewal as worksheet
    dim code, area, yearMonth as Range
    dim rng as Range
    dim ROW_END_DATA, COL_END_DATA, ROW_END_RENEWAL as Long
    dim CODE_COLUMN_RENEWAL, AREA_COLUMN_RENEWAL, YEAR_MONTH_COLUMN_RENEWAL, MAN_COLUMN_RENEWAL As Integer
    dim city, gu, dong as String

    set data = worksheets("data")
    set renewal = worksheets("renewal")

    ROW_END_DATA = data.cells(rows.count, 3).end(xlup).row
    COL_END_DATA = data.cells(3, columns.count).end(Xltoleft).column

    CODE_COLUMN_RENEWAL = 1
    AREA_COLUMN_RENEWAL = 2
    YEAR_MONTH_COLUMN_RENEWAL = 4
    MAN_COLUMN_RENEWAL = 5

    set rng = Range(data.cells(3, 3), data.cells(ROW_END_DATA, COL_END_DATA))

    For Each r In rng

        with renewal

            ROW_END_RENEWAL = .cells(rows.count, 1).end(xlup).row

            set yearMonth = r.end(xlup).end(xlup)
            set code = data.cells(r.row, 1)
            set area = data.cells(r.row, 2)
            city = split(area.value, " ")(0)
            gu = split(area.value, " ")(1)

            If (r.column mod 3 = 0) Then
                .cells(ROW_END_RENEWAL + 1, CODE_COLUMN_RENEWAL) = code.value
                .cells(ROW_END_RENEWAL + 1, AREA_COLUMN_RENEWAL) = city
                .cells(ROW_END_RENEWAL + 1, AREA_COLUMN_RENEWAL + 1) = gu
                .cells(ROW_END_RENEWAL + 1, YEAR_MONTH_COLUMN_RENEWAL) = yearMonth.value
                .cells(ROW_END_RENEWAL + 1, MAN_COLUMN_RENEWAL) = r.value
            else
                .cells(ROW_END_RENEWAL, r.column mod 3 + MAN_COLUMN_RENEWAL) = r.value
            end if
    
        end with

    Next r 

end sub

public sub dataRenewalDong() 

    dim data, renewal as worksheet
    dim code, area, yearMonth as Range
    dim rng as Range
    dim ROW_END_DATA, COL_END_DATA, ROW_END_RENEWAL as Long
    dim CODE_COLUMN_RENEWAL, AREA_COLUMN_RENEWAL, YEAR_MONTH_COLUMN_RENEWAL, MAN_COLUMN_RENEWAL As Integer

    dim city, gu as String

    set data = worksheets("data")
    set renewal = worksheets("renewal")

    ROW_END_DATA = data.cells(rows.count, 3).end(xlup).row
    COL_END_DATA = data.cells(3, columns.count).end(Xltoleft).column

    CODE_COLUMN_RENEWAL = 1
    AREA_COLUMN_RENEWAL = 2
    YEAR_MONTH_COLUMN_RENEWAL = 5
    MAN_COLUMN_RENEWAL = 6

    set rng = Range(data.cells(3, 3), data.cells(ROW_END_DATA, COL_END_DATA))

    For Each r In rng

        with renewal

            ROW_END_RENEWAL = .cells(rows.count, 1).end(xlup).row

            set yearMonth = r.end(xlup).end(xlup)
            set code = data.cells(r.row, 1)
            set area = data.cells(r.row, 2)
            city = split(area.value, " ")(0)
            gu = split(area.value, " ")(1)
            dong = split(area.value, " ")(2)

            If (r.column mod 3 = 0) Then
                .cells(ROW_END_RENEWAL + 1, CODE_COLUMN_RENEWAL) = code.value
                .cells(ROW_END_RENEWAL + 1, AREA_COLUMN_RENEWAL) = city
                .cells(ROW_END_RENEWAL + 1, AREA_COLUMN_RENEWAL + 1) = gu
                .cells(ROW_END_RENEWAL + 1, AREA_COLUMN_RENEWAL + 2) = dong
                .cells(ROW_END_RENEWAL + 1, YEAR_MONTH_COLUMN_RENEWAL) = yearMonth.value
                .cells(ROW_END_RENEWAL + 1, MAN_COLUMN_RENEWAL) = r.value
            else
                .cells(ROW_END_RENEWAL, r.column mod 3 + MAN_COLUMN_RENEWAL) = r.value
            end if
    
        end with

    Next r 

end sub
