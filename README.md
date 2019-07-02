# Get-MSProject-Data
MS Project 2019 Add-in to extract project tasks to a SQL database using EF6
The point of this project is to open a series of MS project files and iterate through each MPP file 
extracting all tasks into a SQL server database.  The tricky part about this code is looping through 
an entity framework DBSET while maintaining state.  I am not including the tables at this time, 
but the tables can be generated from the project's EF edmx. 
