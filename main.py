import sys
from PyQt6.QtWidgets import QApplication
from view.gui import OESAnalyzerGUI

if __name__ == "__main__":
    """
    Main entry point for the OES Analyzer application.
    This script initializes the QApplication and launches the GUI.
    """
    # Initialize the application
    app = QApplication(sys.argv)

    # Create and show the main GUI window
    main_window = OESAnalyzerGUI()
    main_window.show()

    # Execute the application event loop
    sys.exit(app.exec())