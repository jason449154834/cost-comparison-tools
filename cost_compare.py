
from window_control.main_window import *

if __name__ == '__main__':
    app = QApplication([])
    main = main_window()
    main.ui.show()
    app.exec_()

