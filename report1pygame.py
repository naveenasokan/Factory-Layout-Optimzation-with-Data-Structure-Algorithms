import pygame,sys
from time import sleep

def GlobalVariables(module):
    global PATH_COLOR,SOURCE_COLOR,DEST_COLOR,SCREEN_SET_BACK,CONFIG_LAYOUT_GRID_PATH,CONFIG_LAYOUT_PATH,CONFIG_GRID_PATH,IMAGE_PATH
    PATH_COLOR=module.PATH_COLOR
    SOURCE_COLOR=module.SOURCE_COLOR
    DEST_COLOR=module.DEST_COLOR
    SCREEN_SET_BACK=module.SCREEN_SET_BACK
    CONFIG_LAYOUT_GRID_PATH=module.CONFIG_LAYOUT_GRID_PATH
    CONFIG_LAYOUT_PATH=module.CONFIG_LAYOUT_PATH
    CONFIG_GRID_PATH=module.CONFIG_GRID_PATH
    IMAGE_PATH=module.IMAGE_PATH
    
#convert excel colors to RGB
def hexToRgb(value):
    value = value.lstrip('#')
    lv = len(value)
    return tuple(int(value[i:i + lv // 3], 16) for i in range(0, lv, lv // 3))

#draw the entire layout in pygame
def drawLayout(grid):
    for i,row in enumerate(grid):
        for j,cell in enumerate(row):
            if not cell.fill.fgColor.rgb=="00000000":
                rgb_color=hexToRgb(cell.fill.fgColor.rgb)
                pygame.draw.rect(display1,rgb_color[1:], ((j+SCREEN_SET_BACK)*x_width,(i+SCREEN_SET_BACK)*y_width,x_width,y_width))

#fill path colors
def fillPath(end_path_array,distance,color,mech_name,station_name):
    for count,paths in enumerate(end_path_array):
        for path in paths:
            pygame.draw.rect(display1,color, ((path.y+SCREEN_SET_BACK)*x_width,(path.x+SCREEN_SET_BACK)*y_width,x_width,y_width)) #path
        writeText(distance[count],(paths[-1].y+SCREEN_SET_BACK)*x_width,(paths[-1].x+SCREEN_SET_BACK)*y_width,paths)

    writeText1("O",(end_path_array[0][0].y+SCREEN_SET_BACK)*x_width,(end_path_array[0][0].x+SCREEN_SET_BACK)*y_width,(0,0,0))
    pygame.image.save(display1,f"Report1/{mech_name}/{mech_name}_{station_name}.jpg")
    sleep(1)
        #pygame.draw.rect(display1, SOURCE_COLOR, ((paths[0].y+SCREEN_SET_BACK)*x_width,(paths[0].x+SCREEN_SET_BACK)*y_width,x_width,y_width)) #source
        #pygame.draw.rect(display1, DEST_COLOR, ((paths[-1].y+SCREEN_SET_BACK)*x_width,(paths[-1].x+SCREEN_SET_BACK)*y_width,x_width,y_width)) #destination
        
def drawBoundary(ROW,COL):
    x_boundary=x_width*(COL+SCREEN_SET_BACK)
    y_boundary=y_width*(ROW+SCREEN_SET_BACK) 
    pygame.draw.line(display1,(255, 255, 0),(x_boundary,0),(x_boundary,height),3)
    pygame.draw.line(display1,(255, 255, 0),(0,y_boundary),(width,y_boundary),3)
    pygame.draw.line(display1,(255, 255, 0),(SCREEN_SET_BACK*x_width,0),(SCREEN_SET_BACK*x_width,height),3)
    pygame.draw.line(display1,(255, 255, 0),(0,SCREEN_SET_BACK*y_width),(width,SCREEN_SET_BACK*y_width),3)


def writeText1(text,x,y,color):
    font = pygame.font.SysFont('Bold', 25)
    text = font.render(text, True, color)
    display1.blit(text,(x,y))


def writeText(distance,x,y,paths):
    if (paths[-1].x-paths[-2].x)==0:
        if paths[-1].y>paths[-2].y:
            x=x+10
        else:
            x=x-10
    else:
        if paths[-1].x>paths[-2].x:
            y=y+10
        else:
            y=y-10

    font = pygame.font.SysFont('Bold', 20)
    text = font.render(f'{round(distance)}', True, (0,0,0))
    display1.blit(text,(x,y))

def visualize(grid,mech_all,ROW,COL,module):
    GlobalVariables(module)
    global width,height,x_width,y_width,display1
   
   
    pygame.init()
    infoObject = pygame.display.Info()
    image = pygame.image.load(IMAGE_PATH.lstrip("\u202a"))

    #set display
    #width=infoObject.current_w-10 #set entire screen as width for full screen
    #height=infoObject.current_h-90 #set entire screen as height for full screen
    width=1356
    height=678
    display1 = pygame.display.set_mode((width,height))
  
    
    x_width=width//COL #calculate width of each column
    y_width=height//ROW #width of each row
    run=True
    while run:
        display1.fill((255,255,255)) #fill white in background
        if CONFIG_LAYOUT_GRID_PATH:
            display1.blit(image,(SCREEN_SET_BACK*x_width,SCREEN_SET_BACK*y_width))
            drawLayout(grid)
        elif CONFIG_LAYOUT_PATH:
            display1.blit(image,(SCREEN_SET_BACK*x_width,SCREEN_SET_BACK*y_width))
        elif CONFIG_GRID_PATH:
            drawLayout(grid)

        #drawGrids(ROW,COL)
        drawBoundary(ROW,COL)
        
        copy_surface_0=pygame.Surface.copy(display1)
        for mech_class in mech_all:
            for count,station in enumerate(mech_class):
                display1.blit(copy_surface_0,(0,0))
                for element in station.not_reached:
                    writeText1("X",(element.y+SCREEN_SET_BACK)*x_width,(element.x+SCREEN_SET_BACK)*y_width,PATH_COLOR[count])
                copy_surface=pygame.Surface.copy(display1)
                display1.blit(copy_surface,(0,0))
                fillPath(station.end_path_array,station.shortest_dist,PATH_COLOR[count],station.mech_class_name,station.station_name)
        pygame.display.update()
        for event in pygame.event.get():
            print("Done")
            run=False
            pygame.quit()
            break
            #if event.type == pygame.QUIT:
                #pygame.quit()
                #sys.exit()

    pygame.quit()

if __name__=="__main__":
    print("Execute breadthfirstsearch module")


