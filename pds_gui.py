import logging

from pds_generator.requirements_installer import install_missing_requirements


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    install_missing_requirements()
    from pds_generator.gui import PDSGeneratorGUI

    app = PDSGeneratorGUI()
    app.mainloop()
