#FlappyBird Game
import pygame 
import random

pygame.init()

width,height = 500,500
screen = pygame.display.set_mode([width,height])
clock = pygame.time.Clock()

# 1st upper obstacle setting
uobs_x_1 = 150
uobs_y_1 =  0
uobs_h_1 =  200
uobs_w_1 =  70

# 1st lower obstacle setting
lobs_x_1 = uobs_x_1
lobs_y_1 = 400
lobs_h_1 = height
lobs_w_1 = uobs_w_1

# 2nd upper obstacle setting 
uobs_x_2  = 330
uobs_y_2  = uobs_y_1
uobs_h_2  = random.choice([10,20,40,60,120,200,260,310])
uobs_w_2  = uobs_w_1

# 2nd lower obstacle setting
lobs_x_2   = uobs_x_2
lobs_y_2   = random.choice([345,370,400,450,480,350,474,495]) 
lobs_h_2   =  height
lobs_w_2   = uobs_w_2

# player setting
player_x  = 100
player_y  = 300
size      = 35

# main loop
running = True
while running:
    screen.fill([0,0,0])

 # obstacle's movement
    uobs_x_1 -= 1 * 1.11
    lobs_x_1 -= 1 * 1.11
    uobs_x_2 -= 1 * 1.11
    lobs_x_2 -= 1 * 1.11

    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False
        
        # player movement setting    
        if event.type == pygame.KEYDOWN:
            if event.key == pygame.K_UP:
                player_y -= 14 * 0.9
            if event.key == pygame.K_DOWN:
                player_y += 14 * 0.9


    #collision setting with first  obstacle (side by side and from top and bottom)
    if player_x + size > uobs_x_1 and player_x + size < uobs_x_1 + uobs_w_1 and player_x + size > uobs_x_1 and uobs_y_1 < player_y and uobs_y_1 + uobs_h_1 > player_y and uobs_y_1 + uobs_h_1 > player_y + size :
        running = False
 
    if player_y >= uobs_y_1 and uobs_y_1 + uobs_h_1 < player_y + size and uobs_y_1 + uobs_h_1 > player_y and player_x < uobs_x_1 + uobs_w_1 and uobs_x_1 < player_x and uobs_x_1 + uobs_w_1 > player_x + size:
        running = False
    
    if lobs_x_1 < player_x and player_x < lobs_x_1 + lobs_w_1 and player_x + size < lobs_x_1 + lobs_w_1 and lobs_x_1 < player_x + size and player_y < lobs_y_1 and player_y + size > lobs_y_1 and player_y + size < lobs_y_1 + lobs_h_1 and player_y < lobs_y_1 + lobs_h_1:
        running = False
    
    if player_x < lobs_x_1 and player_x + size > lobs_x_1 and player_x < lobs_x_1 + lobs_w_1 and player_x + size < lobs_x_1 + lobs_w_1 and lobs_y_1 < player_y and player_y < lobs_y_1 + lobs_h_1 and player_y + size > lobs_y_1 and player_y +size < lobs_y_1 + lobs_h_1:
        running = False

    if player_x < lobs_x_1 + lobs_w_1 and player_x < lobs_x_1 and player_x + size > lobs_x_1 and player_x + size < lobs_x_1 + lobs_w_1 and player_y < lobs_y_1 and player_y + size > lobs_y_1 and player_y + size < lobs_y_1 +lobs_h_1 and player_y < lobs_y_1 + lobs_h_1:
        running = False

    if player_x < lobs_x_1 + lobs_w_1 and player_x > lobs_x_1 and player_x + size > lobs_x_1 and player_x + size > lobs_x_1 + lobs_w_1 and player_y < lobs_y_1 and player_y + size > lobs_y_1 and player_y + size < lobs_y_1 +lobs_h_1 and player_y < lobs_y_1 + lobs_h_1:
        running = False
    
    if player_x < uobs_x_1 and player_x < uobs_x_1 + uobs_w_1 and uobs_x_1 < player_x + size and player_x + size < uobs_w_1 + uobs_x_1 and uobs_y_1 < player_y and player_y < uobs_y_1 + uobs_h_1 and uobs_y_1 +uobs_h_1 < player_y + size and player_y + size > uobs_y_1:
        running = False

    if player_x > uobs_x_1 and player_x < uobs_x_1 + uobs_w_1 and uobs_x_1 < player_x + size and player_x + size > uobs_w_1 + uobs_x_1 and uobs_y_1 < player_y and player_y < uobs_y_1 + uobs_h_1 and uobs_y_1 +uobs_h_1 < player_y + size and player_y + size > uobs_y_1:
        running = False


   #collision setting with second obstacle (side by side and from top and bottom)

    if player_x + size > uobs_x_2 and player_x + size < uobs_x_2 + uobs_w_2 and player_x + size > uobs_x_2 and uobs_y_2 < player_y and uobs_y_2 + uobs_h_2 > player_y and uobs_y_2 + uobs_h_2 > player_y + size :
        running = False
 
    if player_y >= uobs_y_2 and uobs_y_2 + uobs_h_2 < player_y + size and uobs_y_2 + uobs_h_2 > player_y and player_x < uobs_x_2 + uobs_w_2 and uobs_x_2 < player_x and uobs_x_2 + uobs_w_2 > player_x + size:
        running = False
    
    if lobs_x_2 < player_x and player_x < lobs_x_2 + lobs_w_2 and player_x + size < lobs_x_2 + lobs_w_2 and lobs_x_2 < player_x + size and player_y < lobs_y_2 and player_y + size > lobs_y_2 and player_y + size < lobs_y_2 + lobs_h_2 and player_y < lobs_y_2 + lobs_h_2:
        running = False

    if player_x < lobs_x_2 and player_x + size > lobs_x_2 and player_x < lobs_x_2 + lobs_w_2 and player_x + size < lobs_x_2 + lobs_w_2 and lobs_y_2 < player_y and player_y < lobs_y_2 + lobs_h_2 and player_y + size > lobs_y_2 and player_y +size < lobs_y_2 + lobs_h_2:
        running = False

    if player_x < lobs_x_2 + lobs_w_2 and player_x < lobs_x_2 and player_x + size > lobs_x_2 and player_x + size < lobs_x_2 + lobs_w_2 and player_y < lobs_y_2 and player_y + size > lobs_y_2 and player_y + size < lobs_y_2 +lobs_h_2 and player_y < lobs_y_2 + lobs_h_2:
        running = False

    if player_x < lobs_x_2 + lobs_w_2 and player_x > lobs_x_2 and player_x + size > lobs_x_2 and player_x + size > lobs_x_2 + lobs_w_2 and player_y < lobs_y_2 and player_y + size > lobs_y_2 and player_y + size < lobs_y_2 +lobs_h_2 and player_y < lobs_y_2 + lobs_h_2:
        running = False
    
    if player_x < uobs_x_2 and player_x < uobs_x_2 + uobs_w_2 and uobs_x_2 < player_x + size and player_x + size < uobs_w_2 + uobs_x_2 and uobs_y_2 < player_y and player_y < uobs_y_2 + uobs_h_2 and uobs_y_2 +uobs_h_2 < player_y + size and player_y + size > uobs_y_2:
        running = False

    if player_x > uobs_x_2 and player_x < uobs_x_2 + uobs_w_2 and uobs_x_1 < player_x + size and player_x + size > uobs_w_2 + uobs_x_2 and uobs_y_2 < player_y and player_y < uobs_y_2 + uobs_h_2 and uobs_y_2 +uobs_h_2 < player_y + size and player_y + size > uobs_y_2:
        running = False

    # obstacle respawning after they are offscreen
    if (uobs_x_1 + uobs_w_1)  < 0 and (lobs_x_1 + lobs_w_1) < 0:
        uobs_x_1 = 500
        uobs_h_1 = random.choice([20,100,160,220,270,310,250,260,290,140,120])
        lobs_x_1 = uobs_x_1
        lobs_y_1 = random.choice([400,460,440,480,380,350,370,487,345,450,390])

    if (uobs_x_2 + uobs_w_2)  < 0 and (lobs_x_2 + lobs_w_2) < 0:
        uobs_x_2 = 520
        uobs_h_2  = random.choice([10,20,40,60,120,200,260,310,240,250,280,290,100])
        lobs_x_2 = uobs_x_2
        lobs_y_2 = random.choice([355,370,400,450,480,350,474,495,380,410,430,470,485]) 

 
    # drawing both upper obstacle
    pygame.draw.rect(screen,(0,255,0),(uobs_x_1,uobs_y_1,uobs_w_1,uobs_h_1))
    pygame.draw.rect(screen,(255,0,0),(uobs_x_2,uobs_y_2,uobs_w_2,uobs_h_2))

    # drawing both lower obstacle
    pygame.draw.rect(screen,(0,255,0),(lobs_x_1,lobs_y_1,lobs_w_1,lobs_h_1))
    pygame.draw.rect(screen,(255,0,0),(lobs_x_2,lobs_y_2,lobs_w_2,lobs_h_2))

    # drawing player
    pygame.draw.rect(screen,(0,0,255),(player_x,player_y,size,size))

    pygame.display.update()
    clock.tick(70)
pygame.quit()