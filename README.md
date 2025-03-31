# Python
MyProjects

Printing Business Auto Pricing 2

Auto Pricing 2 is a Python application designed to automate price calculations based on various attributes, like paper type, print type, and color options. It also provides functionalities for saving data into an Excel sheet via a graphical user interface (GUI) built with `tkinter`.

---

## Features

📊 Pricing Automation
- Calculate costs based on paper type (`short`, `long`, `A4`), print type (`text`, `image`, `both`), and color options (`grayscale`, `partial`, `colored`).
- Intuitive GUI for selecting printing options.

🖋️ Excel Integration
- Save printing details into an Excel sheet (`Book1.xlsx`), including:
  - Number of pages
  - Selected attributes (paper type, print type, color type)

🎶 Interactive Features
- Fun audio interaction: a sound effect (`QuackSoundEffect.mp3`) plays when clearing entries.

🖌️ Dynamic Text Fields
- Clear text fields, reset selections, and ensure smooth input without needing to interact directly with the code.

---

Installation

To run the program locally:

1. Clone the repository:
   ```bash
   git clone <repository_url>
   cd AutoPricing2
   ```

2. Install the required Python packages:
   ```bash
   pip install pygame openpyxl
   ```

3. Ensure the following files are available in the project directory:
   - `QuackSoundEffect.mp3` (for audio effects)
   - `Book1.xlsx` (Excel sheet)

4. Run the Python script:
   ```bash
   python auto_pricing_2.py
   ```

---

Usage

1. Launch the application.
2. Select the desired printing options using the interactive buttons.
3. Enter the number of pages to be printed and click the corresponding buttons for paper type, print type, and color options.
4. Save data to the Excel sheet or calculate the total cost.
5. Enjoy the sound effect when resetting inputs!

---

Future Enhancements

- Allow customization of cost values directly through the GUI.
- Add more paper types and print options.
- Enable exporting to additional file formats (CSV, PDF).
- Introduce localization for multilingual support.

---

Dependencies

- `tkinter`: For the graphical user interface.
- `pygame`: For sound effects integration.
- `openpyxl`: For Excel file manipulation.

---

