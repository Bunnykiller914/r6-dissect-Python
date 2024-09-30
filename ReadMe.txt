Attached files:
libr6dissect
libr6dissect.dll
LICENSE
r6-dissect
r6-dissect.exe

These are all used to run the executable to convert the .rec replay files into .json files.

Match Replay Files:
There are a large number of replay files attached to the submission, inside of a .zip folder. Once the program has been run once, folders will initialize in the directory the .py file was run from (this needs to be the same directory as the needed files are in). Put any of these folders into the MatchReplays folder that was created.

The program reads the match replay folder, and looks for a Match folder. Each Match folder contains 1 - 3 Game folders, which each contain up to 15 .rec files. Put the Match folder of choice into the MatchReplays folder. If you would like, you can just extract the .zip into the folder entirely. 

Needed modules:
1. json
2. pandas
3. os
4. openpyxl
5. matplotlib

There is a specific setup for folders, but the program will make them. To start, extract the r6-dissect files into the same directory as the .py file and run the program.
Put any .rec replay files in MatchReplays, and the program will be able to create the .json files and store them in Jsons. After a .json is created, the .rec file will move into MatchReplays\\Dissected. After an excel is created and stored in Outputs, the .json will be moved to Jsons\\Dissected.

Match replays must be stored in a particular way as well. The ones I have sent will be formatted correctly, and the .zip can be extracted into MatchReplays when it is created.

The program can dissect one folder of matches per run, but any amount of .json files can be ran at the same time. The program will combine the results of each of the matches.

The output of the program will be an Excel sheet, named based on the names of the teams in the replay. The workbook will contain an intial sheet with overall stats, and sheets for every player. The player sheets contain the Operators played, and if they won those rounds; and they contain graphs as well. The operators are a WIP.
