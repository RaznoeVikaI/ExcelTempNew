import openpyxl

def readWriteCellFilials():
    wb = openpyxl.load_workbook('задача для кандидата.xlsx')
    # получаем имя активного листа
    active_sheet_name = wb.active
    # print(active_sheet_name)
    # получаем другой лист
    sheet_otchet = wb['отчет']
    print(sheet_otchet)

    # кол-во строки в книге
    rows = sheet_otchet.max_row
    # print(rows)

    # создаем список ключей,чтобы потом перебирать
    # берем этот список из пункта 2
    # пока заглушка
    # keysList = list(volumeIshod.keys())

    # ----------111------------
    # заглушка
    i =1
    print("ПЫТАЕМСЯ ПИСАТЬ")
    indexRowOtchet = 0
    for cell in sheet_otchet['A']:
        indexRowOtchet += 1
        if (indexRowOtchet >= 6):
            cellToInt = int(cell.value)
            # int
            # print(type(cellToInt))
            if (cellToInt == 111):
                if (indexRowOtchet == 7):
                    # for i in keysList:
                        # print("Ключи с расчитаныыми объемами")
                        # print(i)
                        # в зависимости от выбора "номер филиала" расчитывается ячейка "процентные доходы-расходы"
                        # может быть тоже не нужно, т.к. мы ведь берем ячейки уже посчитанные и делаем с ними
                        # вычисления дальше
                        if (i == 1):
                            print("Значения ячейки А i == 1")
                            cellA_Address = "A" + str(indexRowOtchet)
                            print(sheet_otchet[cellA_Address].value)
                            # сюда пишем результат вычисления ячеек E и F
                            cellG_Address = "G" + str(indexRowOtchet)
                            # 3746
                            cellE_Address = "E" + str(indexRowOtchet)
                            # 20,4% -> 0,204
                            cellF_Address = "F" + str(indexRowOtchet)
                            sheet_otchet[cellG_Address].value = \
                                sheet_otchet[cellE_Address].value*(sheet_otchet[cellF_Address].value/100)

    indexRowOtchet += 1
    # Попробуем сохранить файл
    wb.save('задача для кандидата.xlsx')

    # ----------111------------