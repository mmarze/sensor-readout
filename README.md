# sensor-readout
A student project on programming in Python.

_In a mine, measurements are being taken using sensors, the number of which can vary between different measurement series. For each sensor, the results are available in the form of 4 columns: name, description, time, alarm. Each sensor can trigger a warning alarm. Prepare a script to be run from the console, which will read the measurement results from any number of CSV files and present the entirety in the form of a matrix. Examples of input files and the result file can be found in the folder "Exmaple"._

_Requirements:_
1. _Use argparse or similar package_
2. _Input data in the form of several csv files containing n rows and 4 columns: name, description, time, success. The column names may vary, the order will always be the same._
3. _The success of the measurement is determined by its duration. It should be no less than half of the median duration of all measurements. Use the following designations: - (measurement failure), + (measurement success, no alarm), ! (success of measurement, alarm)_
4. _Take into account the occurrence of a different number of sensors over multiple measurements. In the resulting matrix, the absence of a given sensor is marked in the same way as the failure of the measurement._
5. _Allow the choice of presenting the results in the console and/or saving them to an xlsx file, according to the scheme shown in "Example"_
6. _Allow the user to set the date range (names of files to load)_
