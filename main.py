import eel
from handlers_sg import handler
from handlers_sg import handler_tracing
from handlers_sg import weldlog_summary




@eel.expose
def start_handler():
   handler.start_hendler()

@eel.expose
def start_handler_tracing():
   handler_tracing.create_summary_tracing()

@eel.expose
def start_handler_nkdkd():
   weldlog_summary.create_summary_nkdk()



if __name__ == '__main__':
    eel.init('front')
    eel.start('index.html', mode="chrome", size=(700, 400))