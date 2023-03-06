

import openpyxl
import numpy
from openpyxl import load_workbook
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.utils import get_column_letter, column_index_from_string
import time
from collections import deque
import numpy
import re
import copy
import sys
from concurrent.futures import ProcessPoolExecutor as Pool

from pygame.transform import scale
import A_STAR_THREADING


start = time.time()


class MechClass():
    def __init__(self, mech_class_name, station_name, end_path_array, not_reached, shortest_dist, data_reached, data_not_reached, not_available):
        self.station_name = station_name
        self.mech_class_name = mech_class_name
        self.end_path_array = end_path_array
        self.not_reached = not_reached
        self.shortest_dist = shortest_dist
        self.data_reached = data_reached
        self.data_not_reached = data_not_reached
        self.not_available = not_available


def globalVariable2():
    global  EXCEL_PATH 
    EXCEL_PATH = A_STAR_THREADING.EXCEL_PATH


def globalVariable(EXCEL_PATH, mech_name):
    global  REPORT_PATH, START_CELL, END_CELL, SCALE, ROW, COL, grid, book, MECH_NAME
    EXCEL_PATH = A_STAR_THREADING.EXCEL_PATH


    REPORT_PATH = A_STAR_THREADING.REPORT_PATH
    START_CELL = A_STAR_THREADING.START_CELL
    END_CELL = A_STAR_THREADING.END_CELL
    SCALE = A_STAR_THREADING.SCALE
    grid, book = A_STAR_THREADING.readLayoutFromExcel(
        EXCEL_PATH, START_CELL, END_CELL)
    ROW = grid.shape[0]  # total rows
    COL = grid.shape[1]
    MECH_NAME = mech_name


def argumentGenerator(args):
    return A_STAR_THREADING.breadthFirstSearch(*args)


def report1Main(MECH_CLASSES,SOURCE_TEXT_ARRAY):
    mech_all = []
    globalVariable2()
    for MECH_NAME in MECH_CLASSES:
        mech = []
        globalVariable(EXCEL_PATH, MECH_NAME)
        with Pool(3, initializer=globalVariable, initargs=(EXCEL_PATH, MECH_NAME)) as pool:

            results = pool.map(report1MainLogic, SOURCE_TEXT_ARRAY)
        for result in results:
            if result != -1:
                mech = mech+[result]
        mech_all = mech_all+[mech]
    return mech_all, grid, ROW, COL


def report1MainLogic(SOURCE):

    global SEARCH_ARRAY, BLOCK_COLOR, BIN_COLOR
    print(f"Working on {SOURCE} station")
    SOURCE_TEXT = SOURCE
    dest_list_name, mech_name, serial_number = A_STAR_THREADING.readReport2excel(
        REPORT_PATH, MECH_NAME, SOURCE_TEXT)
    dest_list_name = numpy.array(dest_list_name)
    SEARCH_ARRAY, BLOCK_COLOR, BIN_COLOR = A_STAR_THREADING.findColor(book)
    if dest_list_name == -1:
        return -1
    dest_list_name = list(set(dest_list_name[:, 0]))

    not_available = []
    dest_list_address = []
    # Read destination
    for address in dest_list_name:
        source, dest_list_address_temp = A_STAR_THREADING.findVariable(SEARCH_ARRAY,
                                                                       SOURCE_TEXT, [address])  # find variables block color,bin color
        if len(dest_list_address_temp) == 0:
            not_available = not_available + [address]
            continue
        dest_list_address = dest_list_address+dest_list_address_temp
    A_STAR_THREADING.globalVar(BLOCK_COLOR, BIN_COLOR, ROW, COL)
    if not source:
        print("couldnt locate source")
        exit()
    dest_list_address = numpy.array(dest_list_address)
    dest_list_address = list(dest_list_address[:, 0])

    dist_array = []
    end_path_array = []
    data_reached = []
    data_not_reached = []
    not_reached = []
    dummy_data = []
    count = 0
    while count < len(dest_list_address):
        # find source,destination coordinates
        source_rel_x, source_rel_y, dest_rel_x, dest_rel_y = A_STAR_THREADING.findRelativeCoordinates(
            source, dest_list_address[count], START_CELL)

        # point object
        src_pt = A_STAR_THREADING.Point(source_rel_x, source_rel_y)
        dest_pt = A_STAR_THREADING.Point(dest_rel_x, dest_rel_y)

        # check valid source and destination
        check = A_STAR_THREADING.checkSourceDest(src_pt, dest_pt, grid)
        if check == -2:
            print(
                f"{dest_list_name[count]}-{dest_list_address[count]} is out of range")
            not_reached = not_reached+[dest_pt]

            data_not_reached = data_not_reached+[dest_list_name[count]]
            dest_list_name.pop(count)
            dest_list_address.pop(count)

            continue
        elif check == -4:
            print(
                f"{dest_list_name[count]}-{dest_list_address[count]} cant be reached")
            not_reached = not_reached+[dest_pt]
            data_not_reached = data_not_reached+[dest_list_name[count]]
            dest_list_name.pop(count)
            dest_list_address.pop(count)

            continue
        elif check == -1:
            print(f"{SOURCE_TEXT}-{source} is out of range")
            exit()
        elif check == -3:
            print(f"{SOURCE_TEXT}-{source} cant be reached")
            exit()

        count = count+1
        dummy_data = dummy_data+[(src_pt, dest_pt, grid)]

    with Pool(5, initializer=A_STAR_THREADING.globalVar, initargs=(BLOCK_COLOR, BIN_COLOR, ROW, COL)) as pool2:

        result2 = pool2.map(argumentGenerator, dummy_data)

    results2 = list(result2)

    for count, values in enumerate(results2):
        if values[0] >= 0:
            dist_array = dist_array+[values[0]]
            end_path_array = end_path_array+[values[1]]
            data_reached = data_reached + \
                [[dest_list_name[count], round(values[0]*SCALE)]]
        else:
            data_not_reached = data_not_reached+[dest_list_name[count]]
            not_reached = not_reached+[dummy_data[count][1]]
            print(
                f"No path found for {dest_list_name[count]}-{dest_list_address[count]}")

    shortest_dist = list(map(lambda x: round(x*SCALE), dist_array))

    return MechClass(MECH_NAME, SOURCE_TEXT, end_path_array, not_reached, shortest_dist, data_reached, data_not_reached, not_available)
