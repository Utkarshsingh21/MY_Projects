# PONG BATTLE GAME

import pygame
import sys
import random
from pygame import mixer
import os 

# file directory setting
my_dir = "C:\\Users\\UTKARSH SINGH\\AppData\\Local\\Programs\\Python\\Python312\\Games by Utkarsh\\PONG GAME"
os.chdir(my_dir)

# Initialize pygame
pygame.init()

# setting game screen
height = 600
width  = 600
Game_screen = pygame.display.set_mode((width, height))
font = pygame.font.Font(None, 74)

# Game variable
running = True
paused = True
Game_Over = False

# ball
RED = (255, 0, 0)   # ball colour
ball_radius = 15
ball_x = random.choice([50,500])
ball_y = random.choice([100,500])
ball_speed_x = 0.2 #random.choice([0.1,-0.1]) # random left or right
ball_speed_y = -0.2 #random.choice([0.1,-0.1]) # random up or down


# paddle in the bottom corner
White = (255,255,255)  #paddle colour
paddle_width = 100
paddle_height = 20
paddle_x = 250  #bottom-left corner
paddle_y = 600 - paddle_height
paddle_speed = 0.4

# goal area setting
Goal_width = 110
Goal_height = 60
Goal_thickness = 5
show_goal = False
goal_timer = 0

# Score
score = 0
# Main Menu buttons setting
button_width , button_height = 160,60
play_button = pygame.Rect(width // 2 - button_width // 2, height // 2 - 100,button_width,button_height)
pause_button = pygame.Rect(width // 2 - button_width//2 , height // 2,button_width,button_height)
quit_button = pygame.Rect(width // 2 - button_width // 2, height // 2 + 100,button_width,button_height)

# pause_button_b settings
pause_height = 50
pause_width  = 100
pause_button_bx,pause_button_by = 390,5
pause_button_b = pygame.Rect(pause_button_bx,pause_button_by,pause_width,pause_height)

# Quit btn setting
Qbtn_height = 50
Qbtn_width  = 75
Qbtn_x,Qbtn_y = 500,5
Quit_btn = pygame.Rect(Qbtn_x,Qbtn_y,Qbtn_width,Qbtn_height)

# play again btn setting
pab_width = 210
Play_again_btn = pygame.Rect(200,300,pab_width,button_height)

# bckground image setting
bg_image = pygame.image.load("Space-pong-bg.jpg")
bg_image = pygame.transform.scale(bg_image,(600,700))

# background music and sound setting
bg_music = pygame.mixer.music.load("ping-pong-classic-arcade-game.mp3")
pygame.mixer.music.play(-1)
pygame.mixer.music.set_volume(0.5)
# clock for fps 
#clock = pygame.time.Clock()
#fps = 60

# Main Game loop
while running:
    Game_screen.fill(White)
    
    for i in pygame.event.get():
        if i.type == pygame.QUIT:
            pygame.quit()
            sys.exit()
        
        if i.type == pygame.MOUSEBUTTONDOWN:
            if play_button.collidepoint(i.pos):
                paused = False
                Game_Over = False
            elif pause_button.collidepoint(i.pos):
                paused = True
                Game_Over = False
            elif Game_Over and Play_again_btn.collidepoint(i.pos):
                paused = False
                Game_Over = False
                ball_x,ball_y = random.choice([20,500]),random.choice([100,500])
                ball_speed_x , ball_speed_y = 0.2,-0.2
                paddle_x = width // 2 - paddle_width // 2
                score = 0
            elif pause_button_b.collidepoint(i.pos):
                paused = True
                Game_Over = False
            elif quit_button.collidepoint(i.pos) or Quit_btn.collidepoint(i.pos):
                pygame.quit
                sys.exit()
    

    Game_screen.fill((0,0,0))
    Goal_x =  250 #(width - Goal_width) // 2
    Goal_y = 75

    if not paused and not Game_Over:
     #Game_screen.blit(bg_image,(0,75))
     pygame.draw.circle(Game_screen,RED,(ball_x + ball_radius,ball_y + ball_radius),ball_radius)  # drawing ball on screen
     pygame.draw.rect(Game_screen,White,(paddle_x,paddle_y,paddle_width,paddle_height))  # drawing paddle on screen
     pygame.draw.rect(Game_screen,White,(250,75,110,60),5)  #Goal Area
     pygame.draw.line(Game_screen,White,(0,75),(0,600),3)    #left line
     pygame.draw.line(Game_screen,White,(600,75),(600,600),3) # right line
     pygame.draw.line(Game_screen,White,(0,75),(250,75),3)  # line before goal area
     pygame.draw.line(Game_screen,White,(360,75),(600,75),3) # line after goal area

     # drawing buttons on top while game continues
     Mainfont = pygame.font.Font(None, 40)
     pygame.draw.rect(Game_screen,RED,pause_button_b)
     pygame.draw.rect(Game_screen,(255,255,255),Quit_btn)
     pause_text = Mainfont.render("PAUSE",True,White)
     Quit_btn_text = Mainfont.render("QUIT",True,(0,0,0))
     Game_screen.blit(pause_text,(395,15))
     Game_screen.blit(Quit_btn_text,(505,15))


     # Movement of the ball 
     ball_x += ball_speed_x
     ball_y += ball_speed_y

     # Movement of the paddle
     keys = pygame.key.get_pressed()
     if keys[pygame.K_LEFT] and paddle_x > 0:  # Moving the paddle right side and making paddle to not go inside left wall
        paddle_x -= paddle_speed
     if keys[pygame.K_RIGHT] and paddle_x < width - paddle_width:  # Moving the paddle left side and making paddle to not go inside right wall
        paddle_x += paddle_speed

     # ball collides with right and left wall and bouncing back
     if ball_x - ball_radius < 0 or ball_x + 40 > width:      #40 is something common on all wall sides
        ball_speed_x = -ball_speed_x
        

     # ball collides with bottom wall or top wall and bouncing back
     if ball_y - ball_radius < 75:
        ball_speed_y = -ball_speed_y
     if ball_y - ball_radius > (height - 15):  # bottom wall
        Game_Over = True


     # ball collides with paddle and bounce back in random direction
     if(ball_y + ball_radius >= paddle_y and paddle_x <= ball_x <= paddle_x + paddle_width):
        #ball_y = paddle_y - ball_radius # corrects position
        ball_speed_y = -abs(ball_speed_y) # ball bounces back
        collision_sound = mixer.Sound("collision.wav")
        mixer.Sound.play(collision_sound,3)
        mixer.Sound.set_volume(collision_sound,0.5)
        ball_speed_x += random.uniform(-0.1,0.1)  # random adjustments
        ball_speed_x *= 1.1
        ball_speed_y *= 1.1


    # checking if ball is in  goal area 
    if((Goal_x + 10) <= ball_x <= (Goal_x + Goal_width - 10) and Goal_y <= ball_y - ball_radius <= (Goal_y + Goal_height - 30) and not show_goal):
        score += 1
        show_goal = True
        goal_timer = 500
        ball_x,ball_y = random.choice([100,300]),random.choice([150,550]) # after goal it resets the ball position
    
    # displaying "GOAL" on the screen
    if show_goal:
        text = font.render("GOAL!",True,RED)
        Game_screen.blit(text,(width // 2 - 100,height // 2 - 50))
        goal_timer -= 1
        if goal_timer <= 0:
            show_goal = False
    
    # displaying score text
    score_text = font.render(f"Score : {score}",True,White)
    Game_screen.blit(score_text,(10,10))
    
    # drawing play,quit nd pause button
    if paused and not Game_Over:
     Game_screen.blit(bg_image,(0,0))

     # Game Name Text
     Mainfont = pygame.font.Font(None, 100)
     Game_Name_1 = Mainfont.render("PONG BATTLE",True,(0,0,0))
     Game_screen.blit(Game_Name_1,(40,50))
     Game_Name_2 = Mainfont.render("GAME",True,(0,0,0))
     Game_screen.blit(Game_Name_2,(130,110))
     pygame.draw.rect(Game_screen,RED,play_button)
     pygame.draw.rect(Game_screen,RED,pause_button)
     pygame.draw.rect(Game_screen,RED,quit_button)
     
     Mainfont = pygame.font.Font(None, 50)
     play_text = Mainfont.render("PLAY",True,White)
     pause_text = Mainfont.render("PAUSE",True,White)
     quit_text = Mainfont.render("QUIT",True,White)

     Game_screen.blit(play_text,(play_button.x + 40,play_button.y + 15))
     Game_screen.blit(pause_text,(pause_button.x + 20,pause_button.y + 15))
     Game_screen.blit(quit_text,(quit_button.x + 40,quit_button.y + 15))

    if Game_Over is True:
        Game_screen.blit(bg_image,(0,0))

        # displaying Game over Text on Game Over screen
        localfont = pygame.font.Font(None, 120) 
        Game_Over_text = localfont.render("GAME OVER",True,(0,0,0))
        Game_screen.blit(Game_Over_text,(50,150))

        # creating Play again button on game over screen
        localfont = pygame.font.Font(None, 40)
        pygame.draw.rect(Game_screen,(0,0,0),Play_again_btn)
        Play_again_text = localfont.render("PLAY AGAIN",True,White)
        Game_screen.blit(Play_again_text,(225,316))

        # creating quit button on game over screen
        localfont = pygame.font.Font(None, 50)
        pygame.draw.rect(Game_screen,RED,quit_button)
        quit_text = localfont.render("QUIT",True,White)
        Game_screen.blit(quit_text,(260,415))


    pygame.display.flip()    # Update the display screen 
    