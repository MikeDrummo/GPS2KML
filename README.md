# GPS2KML
Waypoints coloured by speed  - convert GPS data to KML via Excel

This macro takes GPS data and produces a KML file which can be imported to Google Earth, for example. The individual waypoints are coloured by speed - red = fast; blue = slow

Data format is
Column 3 = latitude in DD.DDDD    
Column 4 = longitude in DDD.DDDD
Column 5 = Speed
One row for each point

Sub GPSLLSpd2KML(d) does the work and expects a variant array. Different column formats can be easily edited here.
Data is culled to avoid many points that are closer than a specidifed distance
Simple statistics are calc'ed for speed to suggest max and min speeds for colour ranges. Rogue GPS points are often present which skew the fast speeds to too high a value. 
The user picks a folder to save, then a filename.
Finally the file can be opened or dragged into Google Earth
Clicking on a point in GE will show the time and speed for that point.

Sub VCC2KML drives GPSLLSpd2KML by specifying which ranges to copy for Velocitek data. 

for a Selected subset:
Sub VCC2KMLSelection()
Dim d 
  d = Selection.Value
  Call GPSLLSpd2KML(d)
End Sub
