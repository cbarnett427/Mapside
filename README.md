<div align="center">
<h1>Mapside Mapper</h1>
<i>Mapside Mapper is a Python application built for the oil and gas industry for GIS purposes. <br>
It allows users to import coordinates from a CSV or Excel file, calculate distances between consecutive stations, and create KML/KMZ files to visualize lines connecting specific stations in Google Earth Pro.</i>
<br>
<br>

<h2 align="center">
  <img src="https://github.com/cbarnett427/Mapside-Mapper/blob/main/Mapside/assets/ExampleLineGE.png" alt="Example Line Google Earth Image"/>
  <br>
  <br>
<!--   <img src="https://github.com/cbarnett427/Mapside-Mapper/blob/main/Mapside/assets/ExampleLine.png" alt="Example Line Image"/>
  <br>
  <br> -->
  <img src="https://github.com/cbarnett427/Mapside-Mapper/blob/main/Mapside/assets/ExampleGUI.png" width="336" height="240" alt="Example GUI Image"/>
  <br>
  <br>
  <sub><sup>© 2023 Mapside - Licensed under the <a href="./LICENSE">MIT License</a>.</sup></sub>
</h2>
</div>

## Features
- Import coordinates from a CSV or Excel file.
- Calculate distances between consecutive points/stations.
- Adjust the coordinates based on the distances between points/stations.
- Save the adjusted coordinates as a new CSV file named "Mapside Mapping Coordinates Correct.csv."
- Save the original distances in another CSV file named "Mapside Mapping Coordinates Incorrect.csv."
- Generate KML files with lines connecting specific "From Station" and "To Station" coordinates.
- Display tooltips to provide additional information for certain widgets.

## Prerequisites
- Python 3.11 or later
- Required Python libraries (NumPy, Pandas, Tkinter, Simplekml, Geopy, Pyproj, geographiclib)

## How to Use
1. Clone the repository to your local machine or download the zip files and extract the "Mapside" folder to your desktop.
2. Ensure you have the required Python libraries installed by running the following command: <br>pip install numpy pandas tkinter simplekml geopy pyproj geographiclib
3. Make sure to edit the file path variables to your desired file paths.
4. Execute the program by running the "Mapside Mapper.py" file in the "Mapside Program" folder.
5. The application's graphical user interface (GUI) will open. Use the "Import Coordinates" button to select a CSV or Excel file containing the station coordinates.
6. After importing the coordinates, enter the "From Station" and "To Station" numbers in the provided input fields.
7. Click the "Submit" button to generate a KML file with a line connecting the specified "From Station" and "To Station" coordinates.
8. The KML file will be saved on your desktop, and Google Earth Pro will automatically open the file to visualize the line connecting the selected stations.

## Important Notes
- Make sure the station numbers entered in the "From Station" and "To Station" fields are within the valid range. The valid range is displayed on the GUI and depends on the imported coordinate data.
- The total line length should not exceed 65,521 feet. This is the maximum amount of points a KMZ line can have before it does not display correctly in Google Earth Pro.
- The application provides tooltips to help users understand certain features and restrictions.

## License
This project is licensed under the [MIT License](LICENSE).

## Authors
- Clayton Barnett

## Acknowledgments
Special thanks to the creators and maintainers of the libraries used in this project.
