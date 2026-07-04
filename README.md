# Timetable Drafter

This simple script attempts to create a *draft* timetable for university teaching. As an input, it needs a well-structured *Excel* file that contains all the required information and constraints. The output is an *Excel* file with the timetable for various intended audiences.

## Features

* Scheduling multiple programmes.
* Multiple programmes can share modules.
* Multiple academic terms can be scheduled in one go.
* Teaching teams can be accommodated.
* Different types of rooms can be handled.
* Outputs suitable for different audience members: administrator, programme director, module leader, student, etc.

## Notures

This script:
* Cannot schedule timetables with overlapping timeslots.
* Does not do room allocation.
* Does not do fine-tuned adjustments, e.g., if a module is normally taught in seminar rooms, but needs a lab one week, then that is not addressed.

## The Problem

Timetabling is a standard *operations research* problem, usually addressed with the mathematical technique of *linear programming* or *LP* for short. However, most administrators charged with doing timetabling aren't familiar with *LP* and often have to devote large amounts of time and effort on the exercise, which grows exponentially with the number of parameters (programmes, modules, staff members, etc).

This script intends to provide some relief as it undertakes the *LP* and other nifty *operations research* techniques such as Soft Constraints, Orthogonal Bin Packing, Block Rotation, Strict Contiguous Clumping, etc., in the background.

## Usage

### Preparation

Before using this script, ensure you have meticulously completed the setup *Excel* file. There should be 12 sheets in this file (*for ease-of-use, a template Excel file has been provided in this repository*):

#### Configuration

There are a few sheets that provide basic configuration information.

##### Lookup

This is the very last sheet in the workbook and contains several lists. 

***Terms***
This is a list of the academic terms (e.g., semesters, trimesters, etc.). The values should be something like "Autumn" *(there should be no spaces in the term name)*.

***Day***
This is a list of days of the week.

***Timeslot***
This is a list of timeslots for scheduling classes.

IMPORTANT! The script sees these values as mere *lables*, and cannot actually understand the timings. So, a timeslot *8am-10am* is just a phrase to the script and the script does not actually know what to what time it refers or that it is 2 hours long. Therefore, it is essential that the timeslots should not overlap - this script cannot schedule timetables with overlapping timeslots.

***Allocation***
This table has two values *All* and *Split*, and should not be touched. This is used as a lookup for the possible types of staff allocation for modules that are co-taught by multiple staff members.

##### Settings

This is one of the last sheets in the workbook. This contains a settings table with two parameters that take integer input:
| Parameter	| Value |
| --- | --- |
| MaxClassesPerDay | The maximum number of classes a student can be asked to take in a day. |
| MaxStaffClassesPerDay | The maximum number of classes a staff member can be asked to teach in a day. |

If the value field is left empty, the script assumes there is no limit.

##### Module_Cohort_Offset

The script looks at programme cohort sizes to understand how class sizes for modules based on the academic term. However, sometimes, the calculation needs to be offset. For example, let's say that for a programme, the cohort size is 230 in *Autumn* and 135 in *Spring*. Then the modules running in both terms will have 230 students in *Autumn* and 135 in *Spring*, BUT there may be a module that has certain pre-requisites or is scheduled in a way that flips the class size, so it will have 135 students in *Autumn* and 230 in *Spring*. Then for that module, this sheet should be populated as following:
|ModuleCode | Offset |
| --- | --- |
| Module code | 1 |

#### Entities

> ##### Programmes
> 
> This sheet should list the programmes offered.
> | ProgrammeCode | ProgrammeName | AutumnCohort | WinterCohort |
> | --- | --- | --- | --- |
> | The programme code is the unique identifier of the programme. | The full, normal name of the programme. | The number of students in the autumn cohort. | The number of students in the autumn cohort. |
> 
> In this table, the last two columns are illustrative. The column names are created by concatenating term name *(from the Lookup sheet)* and the word "Cohort". At least one term is needed. For additional terms, just add columns to the right and name them precisely - if the term name does not match what's on the *Lookup* sheet exactly, the script will crash.
> 
> ##### Staff
> 
> This sheet should list all staff members engaged in teaching.
> | Initials | FullName |
> | --- | --- |
> | While called initials, this column should have the short unique identifier for the staff member. Sometimes that means staff number. | The name of the staff member. |
> 
> ##### Rooms
> 
> The sheet is called *Rooms*, but the script does not deal with individual rooms. Rather, this sheet should list the different types of rooms.
> | RoomType | Capacity | NumberOfRooms | RoomOverflow |
> | --- | --- | --- | --- |
> | Type of room, e.g. *Seminar*. | The number of students the room can accommodate. | The number of rooms of this type. This is *optional* - if left blank, the script will assume plenty of rooms of this type are available. | This is *optional* - This is intended to build a certain "give" into the room capacity, e.g., if a room's capacity is 30, and the cohort has 31 students, then instead of creating two sections, we may want to subsume the 31st student into the section even though technically the capacity is only 30. This column should contain the amount by which we can expand room capacity *if needed*, e.g., if a 30-student classroom can handle 35, if needed, then this column should say 5. If no value is provided, *zero* is assumed, i.e., room capacity has no "give". |
> 
> ##### Modules
> 
> This sheet should list the different modules offered.
> | ModuleCode | ModuleName | RoomType |
> | --- | --- | --- |
> | The module code is the unique identifier of the module. | The name of the module. | The type of room needed for this module. |
> 
> ##### Timeslots
> 
> This sheet should list all the different timeslots during which teaching can occur.
> | Day | TimeSlot |
> | --- | --- |
> | The day of the week. | The timeslot on this day. The dropdown pulls values from the *Lookup* sheet. |

#### Relationships

> ##### Module_Staff
> 
> This sheet should tell the script who teaches which modules. It can handle teaching teams, i.e., multiple staff members teaching a single module - but each staffer gets a different row. So, if a module is taught by two people, then that module gets two rows - one each for the two staff members.
> | ModuleCode | StaffInitials | AllocationType | SplitPercentage | Term |
> | --- | --- | --- | --- | --- |
> | Module code *(unique identifier)* | Unique identifier of staff member | This is relevant only if multiple staff members are teaching this module. It can take two values: *All* and *Split*. *All* is recommended and means all the staff members will be scheduled for all classes of the module. *Split* is relevant if the module will be broken down into multiple sections, and means that this staff member will teach some of those sections. | This is relevant only if the previous column is *Split* and notes a *rough* proportion of the sections this staff member should be teaching. | The term in which this staff member will teach - this is relevant only for modules that are taught in multiple terms. |
> 
> ##### Module_Programmes
> 
> A module can be offered by multiple programmes and can be taught in multiple terms. This table houses this information.
> | ModuleCode | ProgrammeCode | Term |
> | --- | --- | --- |
> | Module code *(unique identifier)* | Programme code *(unique identifier)* | The term in which the module is offered. |
> 
> If a module is offered by two programmes, then it will need two rows, because the second column takes only one value. If it is offered by three programmes, then it will require three rows, etc.
> If a module is offered in two terms, then it will need two rows, because the second column takes only one value.If it is offered in three terms, then it will require three rows, etc.
> If a module is offered by two programmes and in two terms, then it will require four rows, etc.
> 
> ##### Staff_Availability
> 
> Each staff member will have multiple rows, comprehensively listing all day and timeslot combinations in which they are available to teach.
> | StaffInitials | Day | TimeSlot | Term |
> | --- | --- | --- | --- |
> | Unique identifier of staff member | Day of week available | Timedlot available | This is *optional*. This column caters to the notion that staff availability may be different in different terms. If left blank, it is assumed that the slot indicated in the row applies to all terms. But if a value is provided then the slot provided in the row indicates this staff member's availability only in the indicated term. |
> 
> If a staff member is not listed on this sheet, the script assumes they are always available to teach on any day and any timeslot offered.
> 
> ##### Programme_Availability
> 
> This sheet accounts for the fact that while the institution may run classes on a range of days, the programme may not schedule classes on all the days on which the institution is open, e.g., an EMBA programme may run classes only on Saturdays and Sundays.
> | ProgrammeCode | Day |
> | --- | --- |
> | Programme code *(unique identifier)* | Day on which the programme's classes can be scheduled. |
> 
> Each programme should have one row for each day on which its classes can be scheduled, e.g., if a programme's classes can be scheduled on Monday, Tuesday and Wednesday, then there should be three rows - one each for each day.

### Running the script

When you run the Python script, it will pop up a file explorer/finder window asking you to select the setup *Excel* file.

Once you have done that, it will work through the setup and constraints, develop a timetable, and save the output to an Excel file in the same folder.

### Output

The script will save the output as an *Excel* file in the same folder as the setup file.

The Excel workbook will have a sheet per Word DOCX file and a summary sheet providing an overview for the folder.



### Note

1. The script uses some standard *Python* libraries. If you don't have them installed on your system, then in the first run, the script will try to install these dependencies. If the script can't install these dependencies, for instance if your PC environment precludes it, then it will usually give you the console commands you can use to install these.
2. If some of your libraries and executables sit outside the *Path*, such as if you don't have Admin rights to your work laptop, then you should include the folder addresses in the 'path.txt' file, which should sit in the same folder as the '*timetable-drafter.py*' file, e.g. '*C:\Users\Username\AppData\Roaming\Python\Python313\Scripts*' and '*C:\Users\Username\AppData\Roaming\Python\Python313\site-packages*'.

## Caveat

The script has been written to address *our* situation and requirements. If it can be useful for you, you are welcome to it. However, it will not be suitable for all timetabling situations.

Tested on *Windows 11 Education 64-bit*.

Not tested on *Apple iOS* or *Linux*.

## Never run a Python script before?

It's straightforward, but you may need to install *Python* on your machine first.

### Install Python

*Anaconda* is one of the most popular distributions of *Python*. Download and install from https://www.anaconda.com/download

Installation is simple, but if you need help, check out https://www.anaconda.com/docs/getting-started/anaconda/install/overview

### Start Spyder

*Anaconda* comes with *Spyder IDE*. Start *Spyder*.

Once *Spyder* is ready, open the file '*timetable-drafter.py*' that has the script.

All that's left is for you to hit 'Run', i.e. the green 'Play' button.