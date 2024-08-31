# TEST
# Word Of The Day
import requests
from bs4 import BeautifulSoup
import pandas as pd 

link = 'https://www.dictionary.com/e/word-of-the-day/'
res  = requests.get(link)
soup = BeautifulSoup(res.text,'html.parser')

#Date and Day
date = soup.find('div',class_='otd-item-headword__date').text
date = date.replace(',',' - ')
date = date.replace('\n',' ')

# Word of the day
WOTD      = soup.find('h1',class_='js-fit-text').text
WOTD      = WOTD.upper()

# It gives Part of Speech the following word belongs to.
POTS      = soup.find('span',class_='luna-pos').text
POTS      = POTS.upper()

#Meaning of the following Word
Meaning   = soup.find('div',class_='otd-item-headword__pos').text
Meaning   = (Meaning.split(' '))
Meaning   = (Meaning[2:(len(Meaning)-1)])
Meaning   = ' '.join(Meaning)


# This Whole logic Provides me with use of given word example
Para      = soup.find('div',class_='wotd-item-origin__content')
Examples  = Para.text
Index     = Examples.find(f"EXAMPLES OF {WOTD}")
LEN       = len((f"EXAMPLES OF {WOTD}"))
SIndex    = Index + LEN
Examples  = Examples[(SIndex + 1):(len(Examples)-1)]
Examples  = Examples.replace('\n',' ')


DICT = []   # Creating a empty list.
#df1 = pd.DataFrame(DICT, columns=['Word','POTS','DATE','Meaning','Examples'])   # Creating a Empty DataFrame
df = pd.read_excel("Word_of_The_Day.xlsx")         # Reading the excel sheet rows using for loop
exists = False
for index,row in df.iterrows():
    if row['Word'] == WOTD and row['POTS'] == POTS:
        print('IT Exist')
        exists = True

if not exists:
    print("IT Does not Exist")
    DICT.append([WOTD,POTS,date,Meaning,Examples])

df1 = pd.DataFrame(DICT, columns=['Word','POTS','DATE','Meaning','Examples']) 
# Updating the DataFrame and writing it into the excel sheet
df = pd.concat([df,df1],ignore_index=True)
df.to_excel("Word_of_The_Day.xlsx",index=False) 