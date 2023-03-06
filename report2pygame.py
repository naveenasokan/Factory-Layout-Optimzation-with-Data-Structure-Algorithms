import glob
from time import time
from time import sleep
from PIL import Image
import pygame
import sys
import numpy
import os


def GlobalVariables(module):
    global PATH_COLOR, SOURCE_COLOR, DEST_COLOR, SCREEN_SET_BACK, CONFIG_LAYOUT_GRID_PATH, CONFIG_LAYOUT_PATH, CONFIG_GRID_PATH, IMAGE_PATH, STRIDE
    PATH_COLOR = module.PATH_COLOR
    SOURCE_COLOR = module.SOURCE_COLOR
    DEST_COLOR = module.DEST_COLOR
    SCREEN_SET_BACK = module.SCREEN_SET_BACK
    CONFIG_LAYOUT_GRID_PATH = module.CONFIG_LAYOUT_GRID_PATH
    CONFIG_LAYOUT_PATH = module.CONFIG_LAYOUT_PATH
    CONFIG_GRID_PATH = module.CONFIG_GRID_PATH
    IMAGE_PATH = module.IMAGE_PATH
    STRIDE = module.STRIDE

# convert excel colors to RGB


def hexToRgb(value, cell):
    value = value.lstrip('#')
    lv = len(value)
    return tuple(int(value[i:i + lv // 3], 16) for i in range(0, lv, lv // 3))

# draw the entire layout in pygame


def drawLayout(grid):
    for i, row in enumerate(grid):
        for j, cell in enumerate(row):
            if not cell.fill.fgColor.rgb == "00000000":
                rgb_color = hexToRgb(cell.fill.fgColor.rgb, cell)
                pygame.draw.rect(display1, rgb_color[1:], ((
                    j+SCREEN_SET_BACK)*x_width, (i+SCREEN_SET_BACK)*y_width, x_width, y_width))

# fill path colors


def fillPathAll(end_path_array, distance, copy_surface, station_name, mech_name, path_color, data_all, item_map):
    for count, paths in enumerate(end_path_array):

        for path in paths:
            pygame.draw.rect(display1, path_color, ((path.y+SCREEN_SET_BACK) *
                             x_width, (path.x+SCREEN_SET_BACK)*y_width, x_width, y_width))  # path
        if count < len(end_path_array)-1:
            writeText(str(count+1)+"d", (paths[-1].y+SCREEN_SET_BACK)*x_width,
                      (paths[-1].x+SCREEN_SET_BACK)*y_width, paths, path_color)
            writeText1(count, (paths[0].y+SCREEN_SET_BACK)*x_width,
                       (paths[0].x+SCREEN_SET_BACK)*y_width, paths, path_color)

        else:
            writeText1(count, (paths[0].y+SCREEN_SET_BACK)*x_width,
                       (paths[0].x+SCREEN_SET_BACK)*y_width, paths, path_color)
            writeText(str(0)+"d", (paths[-1].y+SCREEN_SET_BACK)*x_width,
                      (paths[-1].x+SCREEN_SET_BACK)*y_width, paths, path_color)

        writeText2(f"{data_all[count][0]} to {data_all[count][1]}",
                   distance[count], sum(distance[0:count+1]))
        if data_all[count][1] in item_map:
            writematerialname(item_map[data_all[count][1]], "Picked:")
        pygame.image.save(
            display1, f"Report2/{mech_name}/{station_name}/{mech_name}_{station_name} {count+1}.jpg")
        sleep(1)
        display1.blit(copy_surface, (0, 0))

        # pygame.draw.rect(display1, SOURCE_COLOR, ((paths[0].y+SCREEN_SET_BACK)*x_width,(paths[0].x+SCREEN_SET_BACK)*y_width,x_width,y_width)) #source
        # pygame.draw.rect(display1, DEST_COLOR, ((paths[-1].y+SCREEN_SET_BACK)*x_width,(paths[-1].x+SCREEN_SET_BACK)*y_width,x_width,y_width)) #destination


def fillPathOne(end_path_array, distance, copy_surface, start, ss_no, total_distance, station_name, mech_name, path_color, data_one, item_map):

    for count, paths in enumerate(end_path_array):

        for path in paths:
            pygame.draw.rect(display1, path_color, ((path.y+SCREEN_SET_BACK) *
                             x_width, (path.x+SCREEN_SET_BACK)*y_width, x_width, y_width))  # path
        if count % 2 == 0:
            writeText1(0, (paths[0].y+SCREEN_SET_BACK)*x_width,
                       (paths[0].x+SCREEN_SET_BACK)*y_width, paths, path_color)
            writeText(str(start) + "d", (paths[-1].y+SCREEN_SET_BACK)*x_width,
                      (paths[-1].x+SCREEN_SET_BACK)*y_width, paths, path_color)
        else:
            writeText1(start, (paths[0].y+SCREEN_SET_BACK)*x_width,
                       (paths[0].x+SCREEN_SET_BACK)*y_width, paths, path_color)
            writeText(str(0) + "d", (paths[-1].y+SCREEN_SET_BACK)*x_width,
                      (paths[-1].x+SCREEN_SET_BACK)*y_width, paths, path_color)
            start = start+1

        writeText2(f"{data_one[count][0]} to {data_one[count][1]}",
                   distance[count], total_distance+sum(distance[0:count+1]))
        if data_one[count][1] in item_map:
            writematerialname([item_map[data_one[count][1]][0]], "Picked:")
            item_map[data_one[count][1]] = numpy.array(
                item_map[data_one[count][1]])
            item_map[data_one[count][1]] = numpy.roll(
                item_map[data_one[count][1]], -1)

        pygame.image.save(
            display1, f"Report2/{mech_name}/{station_name}/{mech_name}_{station_name} {ss_no +count+1}.jpg")
        sleep(1)
        display1.blit(copy_surface, (0, 0))

    # writeText4(distance,name)


def fillTraffic(end_path_all, end_path_one, distance, station_name, mech_name, path_color, start, data_one, item_map, not_reached):

    for count, paths in enumerate(end_path_all):
        for path in paths:
            pygame.draw.rect(display1, path_color, ((path.y+SCREEN_SET_BACK) *
                             x_width, (path.x+SCREEN_SET_BACK)*y_width, x_width, y_width))  # path
        if count < len(end_path_all)-1:
            writeText(str(count+1), (paths[-1].y+SCREEN_SET_BACK)*x_width,
                      (paths[-1].x+SCREEN_SET_BACK)*y_width, paths, path_color)
        else:
            writeText(str(0), (paths[-1].y+SCREEN_SET_BACK)*x_width,
                      (paths[-1].x+SCREEN_SET_BACK)*y_width, paths, path_color)

    count = 0
    while count < len(end_path_one):
        if count % 2 == 0:
            text = ",".join([str(start+i)
                            for i in range(len(item_map[data_one[count][1]]))])
            for path in end_path_one[count]:
                pygame.draw.rect(display1, path_color, ((
                    path.y+SCREEN_SET_BACK)*x_width, (path.x+SCREEN_SET_BACK)*y_width, x_width, y_width))  # path
            writeText(text, (end_path_one[count][-1].y+SCREEN_SET_BACK)*x_width,
                      (end_path_one[count][-1].x+SCREEN_SET_BACK)*y_width, end_path_one[count], path_color)
            start = start+len(item_map[data_one[count][1]])
            count = count + 2*len(item_map[data_one[count][1]])
    writeText3(f"Distance:{distance}ft  Time:{round((distance/STRIDE)/60,1)}min",
               width/2-120, height-50, (0, 0, 0))

    not_reached = numpy.array(not_reached)
    material_not_picked = []
    if len(not_reached):
        for material_name in not_reached[:, 1]:
            material_not_picked = material_not_picked + \
                list(item_map[material_name])

        material_not_picked = list(set(material_not_picked))
        writematerialname(material_not_picked, "Not Picked:")

    for i in range(1, 6):

        pygame.image.save(
            display1, f"Report2/{mech_name}/{station_name}/{mech_name}_{station_name} all_{i}.jpg")
        sleep(1)
    # writeText4(distance,name)


def drawBoundary(ROW, COL):
    x_boundary = x_width*(COL+SCREEN_SET_BACK)
    y_boundary = y_width*(ROW+SCREEN_SET_BACK)
    pygame.draw.line(display1, (0, 0, 0), (x_boundary, 0),
                     (x_boundary, height), 3)
    pygame.draw.line(display1, (0, 0, 0), (0, y_boundary),
                     (width, y_boundary), 3)
    pygame.draw.line(display1, (0, 0, 0), (SCREEN_SET_BACK *
                     x_width, 0), (SCREEN_SET_BACK*x_width, height), 3)
    pygame.draw.line(display1, (0, 0, 0), (0, SCREEN_SET_BACK *
                     y_width), (width, SCREEN_SET_BACK*y_width), 3)


def drawGrids(ROW, COL):
    for i in range(0, width, x_width):
        if i < (x_width)*COL:
            pygame.draw.line(display1, (200, 200, 200), (i, 0), (i, height))
        else:
            break

    for i in range(0, height, y_width):
        if i < (y_width)*ROW:
            pygame.draw.line(display1, (200, 200, 200), (0, i), (width, i))
        else:
            break
# write notations like 0,1d ,2,2d


def writeText1(count, x, y, paths, color):
    if (paths[0].x-paths[1].x) == 0:
        if paths[0].y > paths[1].y:
            x = x+10
        else:
            x = x-10
    else:
        if paths[0].x > paths[1].x:
            y = y+10
        else:
            y = y-15
    font = pygame.font.SysFont('Bold', 25)
    text = font.render(f'{count}', True, color)
    display1.blit(text, (x, y))


# write notations like 0,1d ,2,2d
def writeText(count, x, y, paths, color):
    if (paths[-1].x-paths[-2].x) == 0:
        if paths[-1].y > paths[-2].y:
            x = x+10
        else:
            x = x-10
    else:
        if paths[-1].x > paths[-2].x:
            y = y+10
        else:
            y = y-15
    font = pygame.font.SysFont('Bold', 25)
    text = font.render(f'{count}', True, color)
    display1.blit(text, (x, y))

# write distance and time taken


def writeText2(path_text, distance, Totaldistance):
    time_taken = round(distance/STRIDE)
    total_time_taken = round((Totaldistance/STRIDE)/60, 1)
    font = pygame.font.SysFont('Bold', 30)
    text1 = font.render(
        f'Distance is {distance} ft , Time taken is {time_taken} sec', True, (0, 0, 0))
    display1.blit(text1, (width/2-120, height-50))
    text2 = font.render(
        f'Total distance is {Totaldistance} ft , Total time taken is {total_time_taken} min', True, (0, 0, 0))
    display1.blit(text2, (width/2-120, height-20))
    text3 = font.render(f'{path_text}', True, (0, 0, 0))
    display1.blit(text3, (width/2-600, height-20))

# Write material picked names


def writematerialname(name, text):
    font = pygame.font.SysFont('Bold', 30)
    x = width-180
    y = height-600
    text1 = font.render(f"{text} ", True, (0, 0, 0))
    display1.blit(text1, (x, y))
    y = y+40
    for i in name:
        text2 = font.render(f"{i} ", True, (0, 0, 0))
        display1.blit(text2, (x, y))
        y = y+40


# Mark not reached with x
def writeText3(text, x, y, path_color):
    font = pygame.font.SysFont('Bold', 30)
    text = font.render(text, True, path_color)
    display1.blit(text, (x, y))

# To get gif image


def GIF(station_name, mech_name):
    # filepaths
    fp_in = f"Report2/{mech_name}/{station_name}/{mech_name}_{station_name}*.jpg"
    fp_out = f"Report2/{mech_name}/{station_name}/{mech_name}_{station_name}.gif"

    img, *imgs = [Image.open(f)
                  for f in sorted(glob.glob(fp_in), key=os.path.getmtime)]
    img.save(fp=fp_out, format='GIF', append_images=imgs,
             save_all=True, duration=1500, loop=0)

    sleep(1)


def visualize(grid, mech_all, ROW, COL, module):
    GlobalVariables(module)
    global width, height, x_width, y_width, display1, y_pos
    pygame.init()
    infoObject = pygame.display.Info()
    image = pygame.image.load(IMAGE_PATH.lstrip("\u202a"))

    # set display
    # width=infoObject.current_w-10 #set entire screen as width for full screen
    # height=infoObject.current_h-90 #set entire screen as height for full screen
    width = 1356
    height = 678
    display1 = pygame.display.set_mode((width, height))

    x_width = width//COL  # calculate width of each column
    y_width = height//ROW  # width of each row
    run = True
    while run:
        display1.fill((255, 255, 255))  # fill white in background
        if CONFIG_LAYOUT_GRID_PATH:
            display1.blit(image, (SCREEN_SET_BACK*x_width,
                          SCREEN_SET_BACK*y_width))
            drawLayout(grid)
        elif CONFIG_LAYOUT_PATH:
            display1.blit(image, (SCREEN_SET_BACK*x_width,
                          SCREEN_SET_BACK*y_width))
        elif CONFIG_GRID_PATH:
            drawLayout(grid)

        drawBoundary(ROW, COL)
        # drawGrids(ROW,COL)
        copy_surface_0 = pygame.Surface.copy(display1)
        for mech_class in mech_all:
            for count, station in enumerate(mech_class):
                display1.blit(copy_surface_0, (0, 0))
                for element in station.not_reached:
                    writeText3("X", (element[0].y+SCREEN_SET_BACK)*x_width,
                               (element[0].x+SCREEN_SET_BACK)*y_width, PATH_COLOR[count])
                copy_surface = pygame.Surface.copy(display1)

                fillPathAll(station.end_path_array_all, station.distance_all, copy_surface, station.station_name,
                            station.mech_class_name, PATH_COLOR[count], station.data_all, station.item_map)
                total_distance_all = sum(station.distance_all)
                if len(station.end_path_array_all) == 0:
                    start = 1
                    ss_no = 0
                else:
                    start = len(station.end_path_array_all)
                    ss_no = len(station.end_path_array_all)
                fillPathOne(station.end_path_array_one, station.distance_one, copy_surface, start, ss_no, total_distance_all,
                            station.station_name, station.mech_class_name, PATH_COLOR[count], station.data_one, station.item_map)
                total_distance = total_distance_all + sum(station.distance_one)
                fillTraffic(station.end_path_array_all, station.end_path_array_one, total_distance, station.station_name,
                            station.mech_class_name, PATH_COLOR[count], start, station.data_one, station.item_map, station.not_reached)

                GIF(station.station_name, station.mech_class_name)

        pygame.display.update()
        for event in pygame.event.get():
            print("Completed")
            run = False
            pygame.quit()
            break
            # if event.type == pygame.QUIT:
            # run=False
            # pygame.quit()
            # sys.exit()

    pygame.quit()


if __name__ == "__main__":
    print("Execute breadthfirstsearch module")
