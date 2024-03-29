![](./img/illu.png)

# traceAreaLas 1.1

traceAreaLas is a small program in Python 3.11 which traces in Autocad the extent of the .las, .laz files.

When you receive a large amount of point clouds from a Lidar, a scouting plan is a definite help.

The clouds containing less than 1000 points are ignored.

Insert a block for each .las file containing

 - 2 Polylines
   - the theoretical area (100m x 100m)
   - the real area in red
   
 - 3 Attributs
   - the name of the file
   - the number of points
   - the date of the survey

![](./img/ilu1.jpg) ![](./img/ilu2.jpg)


## Installation

   - You have to install pyautocad and laspy

```
pip install pyautocad
pip install laspy[lazrs,laszip]
```

## Requirements

  Only for Windows

  Autocad must be installed

## Usage

  - Start Autocad and open the destination drawing.

  - Otherwise Autocad will start with drawing1.dwg.

  - Start traceAreaLas

  - Choose .las or .laz files

  - Wait a minute


## License

  No license
