# Snake Game
import pygame
import sys
import random

# Initialize pygame
pygame.init()

# Main Screen Setting
width = 400
height = 400
size = 20
dir = (size,0)
screen = pygame.display.set_mode((width,height))
clock = pygame.time.Clock()

# Snake's head setting 
snake = [(0,0)]

# food setting
food = (random.randint(0, width//size-1)*size,random.randint(0, height//size-1)*size)

# Main Game loop
while True:
    screen.fill((255,255,255)) 
    
    for i in pygame.event.get():
        if i.type == pygame.QUIT:
            pygame.quit()
            sys.exit()
        
        elif i.type == pygame.KEYDOWN:
            if i.key == pygame.K_UP and dir != (0,size):
                dir = (0,-size)
            
            elif i.key == pygame.K_DOWN and dir != (0,-size):
                dir = (0,size)
            
            elif i.key == pygame.K_LEFT and dir != (size,0):
                dir = (-size,0)
            
            elif i.key == pygame.K_RIGHT and dir != (-size,0):
                dir = (size,0)
    
    # it controls the movement of head and body
    new_head = (snake[0][0] + dir[0], snake[0][1] + dir[1]) # determines the new position for snake's body
    snake.insert(0,new_head) # insert new position in snake for every movement 
# head and body are just simply going moving to next postion according to the list and body is being created and destroyed which is seeming like motion
    x1 = snake[0][0]
    y1 = snake[0][1]
    x2 = food[0]
    y2 = food[1]

    # collision with food
    if(x1 < x2 + size and x1 + size > x2 and y1 < y2 + size and y1 + size > y2):
        food = (random.randint(0, width//size-1)*size,random.randint(0, height//size-1)*size)
    else:
        snake.pop()  # and if body is added without eating the food then that body is popped out 
    
    # collision with walls
    if (x1 < 0 or x1 + size > width or y1 < 0 or y1 + size > height):
        snake = [(0,0)]

    # collision with self body
    l = len(snake)
    if(l > 1):
        for body in snake[1:]:
            if(x1 < body[0] + size and x1 + size > body[0] and y1 < body[1] + size and y1 + size > body[1]):
                snake = [(0,0)]

    # drawing the head and body
    for i in snake:
        pygame.draw.rect(screen,(0,255,0),(i[0],i[1],size,size))

    # drawing the food
    pygame.draw.rect(screen,(255,0,0),(food[0],food[1],size,size))

    # updating the screen
    pygame.display.update()
    clock.tick(5)