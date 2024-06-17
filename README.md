# Automated-Orbital-Data-Processing-and-Analysis
Overview
This project, Automated Orbital Data Processing and Analysis, is designed to process and analyze orbital data from RAR files containing Excel spreadsheets. The application extracts data, processes specific information based on given conditions, and generates visual plots to aid in data analysis. The results include temperature vs. time plots and suggestions based on temperature ranges.

Features
Extract and process multiple RAR files containing Excel files
Filter and analyze data based on specific conditions
Generate temperature vs. time plots
Provide suggestions based on analyzed temperature data
User-friendly GUI for easy navigation and operation
Requirements
Python 3.7+
Tkinter (for GUI)
pandas
matplotlib
rarfile
WinRAR installed on the system
Installation
Clone the Repository

bash
Copy code
git clone https://github.com/yourusername/automated-orbital-data-processing.git
cd automated-orbital-data-processing
Install Dependencies

Make sure you have Python installed on your system. Install the required Python packages:

bash
Copy code
pip install pandas matplotlib rarfile
Ensure WinRAR is Installed

The application requires WinRAR for extracting RAR files. Ensure that UnRAR.exe is available in the default installation path: C:\Program Files\WinRAR\UnRAR.exe.

Usage
Run the Application

bash
Copy code
python main.py
Using the GUI

Select Folder: Choose the folder containing the RAR files.
Select Save Directory: Choose the directory where the plots will be saved.
Process RAR Files: Click the button to start processing the RAR files. The application will display the progress and provide suggestions based on the data analysis.
Output

The application will save temperature vs. time plots in the specified directory.
Suggestions based on the temperature data will be displayed in the GUI.
Code Structure
main.py: The main application script containing the GUI and logic for processing and analyzing the data.
README.md: Project description and usage guide.
Contribution
Feel free to fork this repository, make improvements, and submit pull requests. Contributions are welcome!

License
This project is licensed under the MIT License. See the LICENSE file for details.
