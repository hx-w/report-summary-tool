from ui import *


if __name__ == '__main__':
    global_config.load_config('config.yml')
    create_ui()
    win.mainloop()
