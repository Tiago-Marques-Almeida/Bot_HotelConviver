import schedule
import time
from bot import Bot
import os

def executa():
    Bot.main()

schedule.every().day.at("08:10").do(executa)

while True:
    try:
        schedule.run_pending()
        time.sleep(1)
    except Exception as e:
        with open(os.path.dirname(os.path.dirname(__file__)) + 'log\logerro.txt','w') as arq:
            arq.write(str(e))
