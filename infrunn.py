import pygame
import sys
import os
import random

from pygame.sprite import Group

# initializing pygame and setting screen
pygame.init()
WIDTH, HEIGHT = 1200, 700
screen = pygame.display.set_mode((WIDTH, HEIGHT))
clock = pygame.time.Clock()

# directory of sprites
mydir = "C:\\Users\\UTKARSH SINGH\\AppData\\Local\\Programs\\Python\\Python312\\Games by Utkarsh\\png"
os.chdir(mydir)

#  Colour Constants
WHITE = (255, 255, 255)
GREEN = (0, 255, 0)
RED = (255, 0, 0)


# dino sprites
running_dino = ["Run (1).png","Run (2).png","Run (3).png","Run (4).png","Run (5).png","Run (6).png","Run (7).png","Run (8).png"]
jump_dino = ["Jump (1).png","Jump (2).png","Jump (3).png","Jump (4).png","Jump (5).png","Jump (6).png","Jump (7).png","Jump (8).png","Jump (9).png","Jump (10).png","Jump (11).png","Jump (12).png"]

#collide_dino = ["Dead (1).png","Dead (2).png","Dead (3).png","Dead (4).png","Dead (5).png","Dead (6).png","Dead (7).png","Dead (8).png"]
# obstacles sprites
#game_obstacles = ["obstacle(1).png","obstacle(2).png","obstacle(3).png","obstacle(4).png"]
game_sprites = []  # using for storing resized and transformed sprites temporarly

# function for loading and transforming sprites
def transforming_sprites(sprites):
    game_sprites.clear()
    for i in range(len(sprites)):
        x = sprites[i]
        sprite_image = pygame.image.load(x).convert_alpha()
        sprite_image = pygame.transform.scale(sprite_image,(120,100))
        if isinstance(sprite_image,pygame.Surface):
            game_sprites.append(sprite_image)
    return game_sprites


# storing transformed sprites 

running_dino_sprites = transforming_sprites(running_dino)


# Game Variables
is_jumping = False
is_running = True
gravity = 0.5
jump_height = -10



# Game Variables
anime_time = 100
last_update = pygame.time.get_ticks()

# ground variables
ground_y = 240
jump_dino_sprites = transforming_sprites(jump_dino)


running = True
while running:
    screen.fill(WHITE)
    now = pygame.time.get_ticks()

    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False
            pygame.quit()
            sys.exit()

        if event.type == pygame.KEYDOWN:
            if event.key == pygame.K_SPACE:
                is_running = False
                
                

    # jump settings for dino
    upspeed += gravity
    Diiino.rect.y += Diiino.upspeed 
 

    
    # collision with ground
    if Diiino.rect.y >= ground_y - 90 and Diiino.is_running:
        Diiino.rect.y = ground_y - 90  # Stop at ground
        Diiino.upspeed = 0
        Diiino.is_jumping = False
        #velocity_y = 0  # Stop downward movement
    
    if now - last_update > anime_time and  Diiino.is_jumping and not Diiino.is_running:
        img_inx1 = 0
        img_inx1 = (img_inx1 + 1) % len(jump_dino_sprites)
        last_update = now
        screen.blit(jump_dino_sprites[img_inx1],(10, 150))

    

    # drawing running dino on pygame screen
    #screen.blit(running_dino_sprites[1],(10,150))
   

    
    # drawing obstacales on pygame screen



    # drawing the ground 
    pygame.draw.rect(screen, (0,0,0), (0, ground_y, WIDTH, 5))


    pygame.display.update()
    clock.tick(10)


   

pygame.quit()
"""

        
    
    # moving the obstacles
    ob1_x -= 2
    ob2_x -= 2

    if ob1_x < 0:
        ob1_x = random.randint(0, WIDTH//size-1)*size

    if ob2_x < 0:
        ob2_x = random.randint(137, WIDTH - ob2_wdt)

    if rect_x < ob1_x + ob1_wdt and rect_x + size > ob1_x and rect_y < ob1_y + ob1_hgt and rect_y + size > ob1_y:
        running = False

    if rect_x < ob2_x + ob2_wdt and rect_x + size > ob2_x and rect_y < ob2_y + ob2_hgt and rect_y + size > ob2_y:
        running = False

    if rect_y >= ground_y - 90:
        rect_y = ground_y - 90  # Stop at ground
        velocity_y = 0  # Stop downward movement
        img_inx1 = 0
        #img_inx2 = 0
        is_jumping = False  # Allow another jump
        is_running = True

    if is_jumping:
        is_running = False

    # Drawing and animating player
    jump_frames,run_Frames = jump_animation()
    if now - last_update > anime_time and is_jumping:
        img_inx1 = (img_inx1 + 1) % len(jump_frames)
        last_update = now
        screen.blit(jump_frames[img_inx1],(rect_x, rect_y)) 

    elif now - last_update > anime_time and is_running:
        img_inx2 = (img_inx2 + 1) % len(run_Frames)
        last_update = now
        screen.blit(run_Frames[img_inx2],(rect_x, rect_y))  
    #else:
        #screen.blit(idle_png,(rect_x, rect_y))

    # drawing many obstacle
    if rect_y == ground_y - 90:
        pygame.draw.rect(screen, RED, (ob1_x,ob1_y,ob1_wdt,ob1_hgt))
        pygame.draw.rect(screen, (0,0,255), (ob2_x,ob2_y,ob2_wdt,ob2_hgt))

        class dino(pygame.sprite.Sprite):
    def __init__(self,images,x,y):
        super().__init__()
        self.images = images
        self.image_indx = 0
        self.image = self.images[self.image_indx]
        self.rect = self.image.get_rect(topleft = (x,y))
        self.mask = pygame.mask.from_surface(self.image)
        self.upspeed = 0
        self.is_jumping = False
        self.is_running = True


    def update(self):
        self.image_indx = (self.image_indx + 1) % len(self.images)
        self.image = self.images[self.image_indx]
        self.mask = pygame.mask.from_surface(self.image)

    def jump(self):
        if not self.is_jumping:
            self.upspeed = jump_height
            self.is_jumping = True
            self.is_running = False


class Obstacles(pygame.sprite.Sprite):
    def __init__(self,image,x,y):
        super().__init__()
        self.image = image
        self.rect = self.image.get_rect(topleft = (x,y))
        self.mask = pygame.mask.from_surface(self.image)




Diiino = dino(running_dino_sprites,10,150)




        
# sprite groups
T_rex = pygame.sprite.Group(Diiino)

 new_block = (Oshape[0][0],Oshape[0][1]+20)
    Oshape.insert(0,new_block)
    new_block = (Oshape[0][0]+20,Oshape[0][1])
    Oshape.insert(1,new_block)
    Oshape.pop()
    Oshape.pop()

"""