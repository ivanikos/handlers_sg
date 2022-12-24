import eel
import weldlog_summary
import handler_beta
import handler_tracing




@eel.expose
def start_handler(path):
    try:
       handler_beta.start_handler(path)
       return "Сводка по ФАЗАМ 1, 2, 3 сформирована"
    except:
        return "Возникла ошибка!"
@eel.expose
def start_handler_tracing(path):
    try:
        handler_tracing.create_summary_tracing(path)
        return "Сводка по по теплоспутникам сформирована"
    except:
        return "Возникла ошибка!"


@eel.expose
def start_handler_nkdkd(path):
    try:
        weldlog_summary.create_summary_nkdk(path)
        return "Сводка по % НК СГ сформирована"
    except:
        return "Возникла ошибка!"



if __name__ == '__main__':
    eel.init('front')
    eel.start('index.html', mode="chrome", size=(800, 630))