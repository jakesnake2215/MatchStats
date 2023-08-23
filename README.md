# MatchStats README (Last Changed: 8/23/2023)
A python script to read through json files and show statistics and data from the map.
Working from the work from redraskal on conversion of Match Replay .REC Files to json and xlsx files.
(For greater information on this: https://github.com/redraskal/r6-dissect)
Takes the json file and parses data for a Siege GG like functionality for stat collection.
Also stores this data into an excel file to allow for long term data collection over multiple maps.
Excel will store all player data, as well as individual operator statistics per player (Broken Down by sheets Stats and Operator Stats).

# Basic How To for testing
How to get the json file to use
- First download 'r6-dissect.exe' and store any match replays in same download location
- Open Command Prompt and file path to this location
- refer to redraskal GitHub for greater information (https://github.com/redraskal/r6-dissect)
- Basic Use: type in the command 'r6-dissect -x _MatchJsonName_.json /_MatchReplayFolder_
- In same location will store json file
- In python file find where the json is insert, and insert filepath/filename
- Run Python Script


# To-Do 
Code Refactoring and Restructuring
  - Improve general readability and add greater comments
  - Make the Op Stats Page For Loop print the op names and necessary data points
  - Eliminate Hardcoded Values to make new operator additions more seamless

Bug Fixes/Improvements of Logic
  - Ability to merge names (i.e. when someone changes their name, add ability to combine stats in excel)
  - Improvements of Rating System to better tailor for what are looking for (Can also revert to Reaper's linear approximation of Siege GG rating system [https://www.youtube.com/watch?v=faoQZK2875Q])
  - Look into ability to see defusal of bomb, doesnt seem clear in the json files
  - Fix the Most Played Operator through multiple map, single series logic
    - Needs an error message if not same 10 players

# New Functions to Implement
  - Design and Impliment UI
     - Understand Logic and Flow of User (Can ask a potential user (Whoever at bama does stat tracking))
     - Find and Use a Python Library that can ask user different type of questions (# of maps entered, different options, File Pathing, and show a siege gg like match page like currently seen in python terminal)
     - Find and Use a Python Library to access terminal to run r6-dissect and use json file in Stats script
  - Team Clustering
     - In Excel or other data storage, cluster teams based on what team they play for
     - Create team based stats that make it easier to group players and track map wins and losses
  - Viewing Day Of Stats
     - Temporary File for viewing
     - How to View all Matches (History Tab?)
  - Further Research
     - Greater Usage of Excel Files from Python Scripting
       - Create Functions for excel files to improve readability
     - Changes in the system of 'Stats' excel sheet
