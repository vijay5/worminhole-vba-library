#!/usr/bin/env python
# -*- encoding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      dmnovikov
#
# Created:     10.09.2013
# Copyright:   (c) dmnovikov 2013
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import os
import sys

def main(args):
    print('args =', repr(args))
    if len(args) != 4:
        print('There should be 3 arguments: fileToProcess, textToFind, replaceWith')
        exit()

    fileToProcess = args[1]
    textToFind  = args[2]
    replaceWith = args[3]

    with open (fileToProcess, 'r') as file: # открыли файл для чтения
        fileText = ''
        fileText = file.read()
        fileText = fileText.replace(textToFind, replaceWith)
    # файл закрыт для чтения, в памяти есть текст с заменой символов

    with open (fileToProcess, 'w') as file: # открыли файл для чтения
        file.write(fileText)


if __name__ == '__main__':
    args = sys.argv
    main(args)
else:
    args = []

