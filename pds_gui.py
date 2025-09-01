import logging

from gui import PDSGeneratorGUI


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app = PDSGeneratorGUI()
    app.mainloop()
