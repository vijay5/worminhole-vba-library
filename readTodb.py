#!/usr/bin/env python
# -*- encoding: utf-8 -*-

import sys
import os
import sqlite3

def readToDb(dbPath, tableFilePath, delim=';'):
  chk1 = os.path.isfile(tableFilePath)
  chk2 = os.path.isfile(dbPath)

  if chk1: # есть оба файла
    tableNameArr = os.path.split(tableFilePath)[-1].split('.')
    if len(tableNameArr) > 1:
      tableName = ''.join(tableNameArr[0:-1])
    else:
      tableName = tableNameArr[0]
    tableName = tableName.replace(' ', '_')

    # создаём подключение к БД (даже если файла нет)
    conn = sqlite3.connect(dbPath)
    c = conn.cursor()

    with open(tableFilePath, 'r') as file:
      # первый заход - анализ данных
      rowNum = 0
      colNamesLst = []  # список колонок
      digitSet = set('0123456789.-')
      alphanumSet = set('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_.0123456789-+"\\\/')
      for line in file: # бежим по первым нескольким строкам файла
        rowNum += 1
        lineArr = line.replace('\n','').split(delim)          # "как есть"
        lineArrClr = [el.replace('"','') for el in lineArr] # без кавычек

        if rowNum == 1:
          colNames = dict()
          for el in lineArrClr:
            colName = str(el)
            colName = colName.replace(' ', '_') # замена пробелов в именах колонок
            colNamesLst.append(colName)       # имена столбцов по-порядку
            if colNames.get(colName) == None: # если нет колонки с таким именем
              colNames[colName] = dict() # создаём с пустым словарём типов значений

        if rowNum > 1: # для второй и последующих строк
          cnt1 = -1
          for el in lineArr: # перебор элеметов текущей строки
            cnt1 += 1
            key = str(el) # явное приведение к тексту
            colName = colNamesLst[cnt1]
            tmpDict = colNames[colName]
            # пытаемся определить тип данных
            if key != '' and set(key).issubset(digitSet):
              if '.' in key:
                typeName = 'real'
              else:
                typeName = 'integer'
            elif '"' in key:
              typeName = 'text'
            elif ':' in key:
              typeName = 'date'
            elif '"' in key:
              typeName = 'text'
            elif key != '' and set(key).issubset(alphanumSet):
              typeName = 'text'
            else: # для остальных типов
              typeName = 'none'

            if tmpDict.get(typeName): # есть такой тип в словаре
              tmpDict[typeName] = tmpDict[typeName] + 1
            else: # такого типа нет - создаём
              tmpDict[typeName] = 1

            colNames[colName] = tmpDict # пишем словарь обратно

        if rowNum > 300: # для определения формата данных используем только первые 20 строк
          break


    # в этой точке знаем тип данных - осталось выбрать самый часто встречающийся
    tableFields = []       # список полей таблицы с типами
    colTypes = []          # типы данных для будущего конвертора
    correctedColTypes = [] # типы данных для строки создания таблицы
    for colName in colNamesLst:
      tmpDict = colNames[colName] # словарь с частотами
      maxFreq = 0
      maxTypeName = ''
      for (typeName, freq) in tmpDict.items():
        if freq > maxFreq:
          maxFreq = freq
          maxTypeName = typeName
      # в этой точке известен самый часто встречающийся тип
      if 'text' in tmpDict.keys():
        maxTypeName = 'text'
      colTypes.append(maxTypeName)

      if maxTypeName == 'date':
        maxTypeName = 'text' # теоретически, можно навесить конвертор
      elif maxTypeName == 'none':
        maxTypeName = 'text'
      else:
        pass

      correctedColTypes.append(maxTypeName) # добавляем "исправленное" поле
      tableFields.append('[' + colName + '] ' + maxTypeName) # добавляем поле с описанием

    # создаём таблицу, если её не было
    c.execute('create table if not exists ' + tableName + ' (' + ', '.join(tableFields) + ')')
    conn.commit()
    c.execute('delete from ' + tableName) # чистим данные в таблице
    conn.commit()
    # в этой точке существует или создана таблица

    # читаем остальные строки в БД
    # commit каждые n записей + после прочтения всех строк (контрольный)
    with open(tableFilePath, 'r') as file:
      # первый заход - анализ данных
      rowNum = 0
      qnMarks = ', '.join(['?' for i in range(0, len(colNamesLst))]) # для insert'а
      recordList = []

      for line in file: # цикл по файлу
        rowNum += 1
        lineArr = line.replace('\n','').split(delim)          # "как есть"
        lineArrClr = [el.replace('"','') for el in lineArr] # без кавычек

        if len(lineArr) != len(colNamesLst):
          print('Строка', rowNum, 'содержит неправильное число полей')

        if rowNum > 1 and len(lineArr) == len(colNamesLst):
          for colNum in range(0, len(lineArrClr)): # перебор номеров столбцов
            value = lineArrClr[colNum] # значение в текущей строке
            if colTypes[colNum] in ('date', 'text', 'none'):
              lineArrClr[colNum] = str(value)
            elif colTypes[colNum] == 'real':
              if value == '' or value == None:
                lineArrClr[colNum] = 0
              else:
                lineArrClr[colNum] = float(value)
            elif colTypes[colNum] == 'integer':
              if value == '' or value == None:
                lineArrClr[colNum] = 0
              else:
                lineArrClr[colNum] = int(value)

          record = tuple(lineArrClr)
          recordList.append(record)
          if rowNum % 50 == 0: # с шагом 10 строк
            c.executemany('insert into ' + tableName + ' (' + ', '.join(colNamesLst) + ') values ('+ qnMarks +')', recordList)
            recordList = [] # чистим буфер
#            conn.commit()

        if rowNum % 500 == 0:
          #conn.commit()
          pass
        if rowNum % 100000 == 0:
          # conn.commit()
          # break # выходим из цикла
          print('Обработано', str(rowNum), 'строк')
      if len(recordList) > 0:
        c.executemany('insert into ' + tableName + ' (' + ', '.join(colNamesLst) + ') values ('+ qnMarks +')', recordList)

      conn.commit() # финальный (на всякий пожарный)
      print('_______________________________________________')
      print('Всего обработано', str(rowNum), 'строк')

  else:
    print('Не удалось открыть файл или БД')


if __name__ != '__main__':
  args = sys.argv
  if len(args) == 3:
    dbPath = args[1]
    tableFilePath  = args[2]
    readToDb(dbPath, tableFilePath)
  elif len(args) >= 4:
    dbPath = args[1]
    tableFilePath  = args[2]
    colDelim  = args[3]
    readToDb(dbPath, tableFilePath, colDelim)
  else:
    print('Задано неверное количество аргументов')
else:
  dbPath = 'D:\\Users\\User\\Desktop\\Текучка\\(xxxx)+ Чтение мегафайла для Сергея\\temp.db'
  tableFilePath  = 'D:\\Users\\User\\Desktop\\Текучка\\(xxxx)+ Чтение мегафайла для Сергея\\tshist.csv'
  readToDb(dbPath, tableFilePath, ',')



