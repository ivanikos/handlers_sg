import eel




@eel.expose
# defining the function for addition of two numbers
def add(data_1, data_2):
    int1 = int(data_1)
    int2 = int(data_2)
    output = int1 + int2
    return output

if __name__ == '__main__':
    eel.init('front')
    eel.start('index.html', mode="chrome", size=(700, 400))