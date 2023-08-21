import json
import math
import numpy as np
import statistics
import openpyxl
from openpyxl import load_workbook


operatorsToNum = {"Mute": 0, "Smoke": 1, "Castle": 2, "Pulse": 3, "Doc": 4, "Rook": 5, "Jager": 6, "Bandit": 7,"Tachanka": 8, "Kapkan": 9, "Frost": 10, "Valkyrie": 11, "Caveira": 12, "Echo": 13, "Mira": 14,"Lesion": 15, "Ela": 16, "Vigil": 17, "Alibi": 18, "Maestro": 19, "Clash": 20, "Kaid": 21, "Mozzie": 22,"Warden": 23, "Goyo": 24, "Wamai": 25, "Oryx": 26, "Melusi": 27, "Aruni": 28, "Thunderbird": 29,"Thorn": 30, "Azami": 31, "Solis": 32, "Fenrir": 33, "Sledge": 34, "Thatcher": 35, "Ash": 36,"Thermite": 37, "Montagne": 38, "Twitch": 39, "Blitz": 40, "IQ": 41, "Fuze": 42, "Glaz": 43,"Buck": 44, "Blackbeard": 45, "Capitao": 46, "Hibana": 47, "Jackal": 48, "Ying": 49, "Zofia": 50,"Dokkaebi": 51, "Finka": 52, "Lion": 53, "Maverick": 54, "Nomad": 55, "Gridlock": 56, "Nokk": 57,"Amaru": 58, "Kali": 59, "Iana": 60, "Ace": 61, "Zero": 62, "Flores": 63, "Osa": 64, "Sens": 65,"Grim": 66, "Brava": 67}
numToOperators = {0: 'Mute', 1: 'Smoke', 2: 'Castle', 3: 'Pulse', 4: 'Doc', 5: 'Rook', 6: 'Jager', 7: 'Bandit', 8: 'Tachanka', 9: 'Kapkan', 10: 'Frost', 11: 'Valkyrie', 12: 'Caveira', 13: 'Echo', 14: 'Mira', 15: 'Lesion', 16: 'Ela', 17: 'Vigil', 18: 'Alibi', 19: 'Maestro', 20: 'Clash', 21: 'Kaid', 22: 'Mozzie', 23: 'Warden', 24: 'Goyo', 25: 'Wamai', 26: 'Oryx', 27: 'Melusi', 28: 'Aruni', 29: 'Thunderbird', 30: 'Thorn', 31: 'Azami', 32: 'Solis', 33: 'Fenrir', 34: 'Sledge', 35: 'Thatcher', 36: 'Ash', 37: 'Thermite', 38: 'Montagne', 39: 'Twitch', 40: 'Blitz', 41: 'IQ', 42: 'Fuze', 43: 'Glaz', 44: 'Buck', 45: 'Blackbeard', 46: 'Capitao', 47: 'Hibana', 48: 'Jackal', 49: 'Ying', 50: 'Zofia', 51: 'Dokkaebi', 52: 'Finka', 53: 'Lion', 54: 'Maverick', 55: 'Nomad', 56: 'Gridlock', 57: 'Nokk', 58: 'Amaru', 59: 'Kali', 60: 'Iana', 61: 'Ace', 62: 'Zero', 63: 'Flores', 64: 'Osa', 65: 'Sens', 66: 'Grim', 67: 'Brava'}
#File Path Location of Excel Sheet
excelFile = 'C:\\Users\\jakeg\\OneDrive\\Desktop\\r6-dissect-v0.11.1-windows-amd64\\Stats.xlsx'
#Load the workbook into read cells
workbook = load_workbook(filename=excelFile)
timeToTrade = 10 #seconds
opNumbers = 68 #number of operators
# <= 33 is defender, > 33 is attacker
#Header of 'Operator Stats' page in Excel
operatorHeader = ['Username', 'Most Played Atk Op', '# of Rounds', 'Most Played Def Op', '# of Rounds','Highest Rating Atk Op', 'Rating', 'Highest Rating Def Op', 'Rating','Mute Rating', 'Mute Kills', 'Mute Deaths', 'Mute EntryKills', 'Mute EntryDeaths', 'Mute KOST', 'Mute HS', 'Mute MKills', 'Mute Trades', 'Mute Clutch', 'Mute Plants', 'Mute Rounds', 'Smoke Rating', 'Smoke Kills', 'Smoke Deaths', 'Smoke EntryKills', 'Smoke EntryDeaths', 'Smoke KOST', 'Smoke HS', 'Smoke MKills', 'Smoke Trades', 'Smoke Clutch', 'Smoke Plants', 'Smoke Rounds', 'Castle Rating', 'Castle Kills', 'Castle Deaths', 'Castle EntryKills', 'Castle EntryDeaths', 'Castle KOST', 'Castle HS', 'Castle MKills', 'Castle Trades', 'Castle Clutch', 'Castle Plants', 'Castle Rounds', 'Pulse Rating', 'Pulse Kills', 'Pulse Deaths', 'Pulse EntryKills', 'Pulse EntryDeaths', 'Pulse KOST', 'Pulse HS', 'Pulse MKills', 'Pulse Trades', 'Pulse Clutch', 'Pulse Plants', 'Pulse Rounds', 'Doc Rating', 'Doc Kills', 'Doc Deaths', 'Doc EntryKills', 'Doc EntryDeaths', 'Doc KOST', 'Doc HS', 'Doc MKills', 'Doc Trades', 'Doc Clutch', 'Doc Plants', 'Doc Rounds', 'Rook Rating', 'Rook Kills', 'Rook Deaths', 'Rook EntryKills', 'Rook EntryDeaths', 'Rook KOST', 'Rook HS', 'Rook MKills', 'Rook Trades', 'Rook Clutch', 'Rook Plants', 'Rook Rounds', 'Jager Rating', 'Jager Kills', 'Jager Deaths', 'Jager EntryKills', 'Jager EntryDeaths', 'Jager KOST', 'Jager HS', 'Jager MKills', 'Jager Trades', 'Jager Clutch', 'Jager Plants', 'Jager Rounds', 'Bandit Rating', 'Bandit Kills', 'Bandit Deaths', 'Bandit EntryKills', 'Bandit EntryDeaths', 'Bandit KOST', 'Bandit HS', 'Bandit MKills', 'Bandit Trades', 'Bandit Clutch', 'Bandit Plants', 'Bandit Rounds', 'Tachanka Rating', 'Tachanka Kills', 'Tachanka Deaths', 'Tachanka EntryKills', 'Tachanka EntryDeaths', 'Tachanka KOST', 'Tachanka HS', 'Tachanka MKills', 'Tachanka Trades', 'Tachanka Clutch', 'Tachanka Plants', 'Tachanka Rounds', 'Kapkan Rating', 'Kapkan Kills', 'Kapkan Deaths', 'Kapkan EntryKills', 'Kapkan EntryDeaths', 'Kapkan KOST', 'Kapkan HS', 'Kapkan MKills', 'Kapkan Trades', 'Kapkan Clutch', 'Kapkan Plants', 'Kapkan Rounds', 'Frost Rating', 'Frost Kills', 'Frost Deaths', 'Frost EntryKills', 'Frost EntryDeaths', 'Frost KOST', 'Frost HS', 'Frost MKills', 'Frost Trades', 'Frost Clutch', 'Frost Plants', 'Frost Rounds', 'Valkyrie Rating', 'Valkyrie Kills', 'Valkyrie Deaths', 'Valkyrie EntryKills', 'Valkyrie EntryDeaths', 'Valkyrie KOST', 'Valkyrie HS', 'Valkyrie MKills', 'Valkyrie Trades', 'Valkyrie Clutch', 'Valkyrie Plants', 'Valkyrie Rounds', 'Caveira Rating', 'Caveira Kills', 'Caveira Deaths', 'Caveira EntryKills', 'Caveira EntryDeaths', 'Caveira KOST', 'Caveira HS', 'Caveira MKills', 'Caveira Trades', 'Caveira Clutch', 'Caveira Plants', 'Caveira Rounds', 'Echo Rating', 'Echo Kills', 'Echo Deaths', 'Echo EntryKills', 'Echo EntryDeaths', 'Echo KOST', 'Echo HS', 'Echo MKills', 'Echo Trades', 'Echo Clutch', 'Echo Plants', 'Echo Rounds', 'Mira Rating', 'Mira Kills', 'Mira Deaths', 'Mira EntryKills', 'Mira EntryDeaths', 'Mira KOST', 'Mira HS', 'Mira MKills', 'Mira Trades', 'Mira Clutch', 'Mira Plants', 'Mira Rounds', 'Lesion Rating', 'Lesion Kills', 'Lesion Deaths', 'Lesion EntryKills', 'Lesion EntryDeaths', 'Lesion KOST', 'Lesion HS', 'Lesion MKills', 'Lesion Trades', 'Lesion Clutch', 'Lesion Plants', 'Lesion Rounds', 'Ela Rating', 'Ela Kills', 'Ela Deaths', 'Ela EntryKills', 'Ela EntryDeaths', 'Ela KOST', 'Ela HS', 'Ela MKills', 'Ela Trades', 'Ela Clutch', 'Ela Plants', 'Ela Rounds', 'Vigil Rating', 'Vigil Kills', 'Vigil Deaths', 'Vigil EntryKills', 'Vigil EntryDeaths', 'Vigil KOST', 'Vigil HS', 'Vigil MKills', 'Vigil Trades', 'Vigil Clutch', 'Vigil Plants', 'Vigil Rounds', 'Alibi Rating', 'Alibi Kills', 'Alibi Deaths', 'Alibi EntryKills', 'Alibi EntryDeaths', 'Alibi KOST', 'Alibi HS', 'Alibi MKills', 'Alibi Trades', 'Alibi Clutch', 'Alibi Plants', 'Alibi Rounds', 'Maestro Rating', 'Maestro Kills', 'Maestro Deaths', 'Maestro EntryKills', 'Maestro EntryDeaths', 'Maestro KOST', 'Maestro HS', 'Maestro MKills', 'Maestro Trades', 'Maestro Clutch', 'Maestro Plants', 'Maestro Rounds', 'Clash Rating', 'Clash Kills', 'Clash Deaths', 'Clash EntryKills', 'Clash EntryDeaths', 'Clash KOST', 'Clash HS', 'Clash MKills', 'Clash Trades', 'Clash Clutch', 'Clash Plants', 'Clash Rounds', 'Kaid Rating', 'Kaid Kills', 'Kaid Deaths', 'Kaid EntryKills', 'Kaid EntryDeaths', 'Kaid KOST', 'Kaid HS', 'Kaid MKills', 'Kaid Trades', 'Kaid Clutch', 'Kaid Plants', 'Kaid Rounds', 'Mozzie Rating', 'Mozzie Kills', 'Mozzie Deaths', 'Mozzie EntryKills', 'Mozzie EntryDeaths', 'Mozzie KOST', 'Mozzie HS', 'Mozzie MKills', 'Mozzie Trades', 'Mozzie Clutch', 'Mozzie Plants', 'Mozzie Rounds', 'Warden Rating', 'Warden Kills', 'Warden Deaths', 'Warden EntryKills', 'Warden EntryDeaths', 'Warden KOST', 'Warden HS', 'Warden MKills', 'Warden Trades', 'Warden Clutch', 'Warden Plants', 'Warden Rounds', 'Goyo Rating', 'Goyo Kills', 'Goyo Deaths', 'Goyo EntryKills', 'Goyo EntryDeaths', 'Goyo KOST', 'Goyo HS', 'Goyo MKills', 'Goyo Trades', 'Goyo Clutch', 'Goyo Plants', 'Goyo Rounds', 'Wamai Rating', 'Wamai Kills', 'Wamai Deaths', 'Wamai EntryKills', 'Wamai EntryDeaths', 'Wamai KOST', 'Wamai HS', 'Wamai MKills', 'Wamai Trades', 'Wamai Clutch', 'Wamai Plants', 'Wamai Rounds', 'Oryx Rating', 'Oryx Kills', 'Oryx Deaths', 'Oryx EntryKills', 'Oryx EntryDeaths', 'Oryx KOST', 'Oryx HS', 'Oryx MKills', 'Oryx Trades', 'Oryx Clutch', 'Oryx Plants', 'Oryx Rounds', 'Melusi Rating', 'Melusi Kills', 'Melusi Deaths', 'Melusi EntryKills', 'Melusi EntryDeaths', 'Melusi KOST', 'Melusi HS', 'Melusi MKills', 'Melusi Trades', 'Melusi Clutch', 'Melusi Plants', 'Melusi Rounds', 'Aruni Rating', 'Aruni Kills', 'Aruni Deaths', 'Aruni EntryKills', 'Aruni EntryDeaths', 'Aruni KOST', 'Aruni HS', 'Aruni MKills', 'Aruni Trades', 'Aruni Clutch', 'Aruni Plants', 'Aruni Rounds', 'Thunderbird Rating', 'Thunderbird Kills', 'Thunderbird Deaths', 'Thunderbird EntryKills', 'Thunderbird EntryDeaths', 'Thunderbird KOST', 'Thunderbird HS', 'Thunderbird MKills', 'Thunderbird Trades', 'Thunderbird Clutch', 'Thunderbird Plants', 'Thunderbird Rounds', 'Thorn Rating', 'Thorn Kills', 'Thorn Deaths', 'Thorn EntryKills', 'Thorn EntryDeaths', 'Thorn KOST', 'Thorn HS', 'Thorn MKills', 'Thorn Trades', 'Thorn Clutch', 'Thorn Plants', 'Thorn Rounds', 'Azami Rating', 'Azami Kills', 'Azami Deaths', 'Azami EntryKills', 'Azami EntryDeaths', 'Azami KOST', 'Azami HS', 'Azami MKills', 'Azami Trades', 'Azami Clutch', 'Azami Plants', 'Azami Rounds', 'Solis Rating', 'Solis Kills', 'Solis Deaths', 'Solis EntryKills', 'Solis EntryDeaths', 'Solis KOST', 'Solis HS', 'Solis MKills', 'Solis Trades', 'Solis Clutch', 'Solis Plants', 'Solis Rounds', 'Fenrir Rating', 'Fenrir Kills', 'Fenrir Deaths', 'Fenrir EntryKills', 'Fenrir EntryDeaths', 'Fenrir KOST', 'Fenrir HS', 'Fenrir MKills', 'Fenrir Trades', 'Fenrir Clutch', 'Fenrir Plants', 'Fenrir Rounds', 'Sledge Rating', 'Sledge Kills', 'Sledge Deaths', 'Sledge EntryKills', 'Sledge EntryDeaths', 'Sledge KOST', 'Sledge HS', 'Sledge MKills', 'Sledge Trades', 'Sledge Clutch', 'Sledge Plants', 'Sledge Rounds', 'Thatcher Rating', 'Thatcher Kills', 'Thatcher Deaths', 'Thatcher EntryKills', 'Thatcher EntryDeaths', 'Thatcher KOST', 'Thatcher HS', 'Thatcher MKills', 'Thatcher Trades', 'Thatcher Clutch', 'Thatcher Plants', 'Thatcher Rounds', 'Ash Rating', 'Ash Kills', 'Ash Deaths', 'Ash EntryKills', 'Ash EntryDeaths', 'Ash KOST', 'Ash HS', 'Ash MKills', 'Ash Trades', 'Ash Clutch', 'Ash Plants', 'Ash Rounds', 'Thermite Rating', 'Thermite Kills', 'Thermite Deaths', 'Thermite EntryKills', 'Thermite EntryDeaths', 'Thermite KOST', 'Thermite HS', 'Thermite MKills', 'Thermite Trades', 'Thermite Clutch', 'Thermite Plants', 'Thermite Rounds', 'Montagne Rating', 'Montagne Kills', 'Montagne Deaths', 'Montagne EntryKills', 'Montagne EntryDeaths', 'Montagne KOST', 'Montagne HS', 'Montagne MKills', 'Montagne Trades', 'Montagne Clutch', 'Montagne Plants', 'Montagne Rounds', 'Twitch Rating', 'Twitch Kills', 'Twitch Deaths', 'Twitch EntryKills', 'Twitch EntryDeaths', 'Twitch KOST', 'Twitch HS', 'Twitch MKills', 'Twitch Trades', 'Twitch Clutch', 'Twitch Plants', 'Twitch Rounds', 'Blitz Rating', 'Blitz Kills', 'Blitz Deaths', 'Blitz EntryKills', 'Blitz EntryDeaths', 'Blitz KOST', 'Blitz HS', 'Blitz MKills', 'Blitz Trades', 'Blitz Clutch', 'Blitz Plants', 'Blitz Rounds', 'IQ Rating', 'IQ Kills', 'IQ Deaths', 'IQ EntryKills', 'IQ EntryDeaths', 'IQ KOST', 'IQ HS', 'IQ MKills', 'IQ Trades', 'IQ Clutch', 'IQ Plants', 'IQ Rounds', 'Fuze Rating', 'Fuze Kills', 'Fuze Deaths', 'Fuze EntryKills', 'Fuze EntryDeaths', 'Fuze KOST', 'Fuze HS', 'Fuze MKills', 'Fuze Trades', 'Fuze Clutch', 'Fuze Plants', 'Fuze Rounds', 'Glaz Rating', 'Glaz Kills', 'Glaz Deaths', 'Glaz EntryKills', 'Glaz EntryDeaths', 'Glaz KOST', 'Glaz HS', 'Glaz MKills', 'Glaz Trades', 'Glaz Clutch', 'Glaz Plants', 'Glaz Rounds', 'Buck Rating', 'Buck Kills', 'Buck Deaths', 'Buck EntryKills', 'Buck EntryDeaths', 'Buck KOST', 'Buck HS', 'Buck MKills', 'Buck Trades', 'Buck Clutch', 'Buck Plants', 'Buck Rounds', 'Blackbeard Rating', 'Blackbeard Kills', 'Blackbeard Deaths', 'Blackbeard EntryKills', 'Blackbeard EntryDeaths', 'Blackbeard KOST', 'Blackbeard HS', 'Blackbeard MKills', 'Blackbeard Trades', 'Blackbeard Clutch', 'Blackbeard Plants', 'Blackbeard Rounds', 'Capitao Rating', 'Capitao Kills', 'Capitao Deaths', 'Capitao EntryKills', 'Capitao EntryDeaths', 'Capitao KOST', 'Capitao HS', 'Capitao MKills', 'Capitao Trades', 'Capitao Clutch', 'Capitao Plants', 'Capitao Rounds', 'Hibana Rating', 'Hibana Kills', 'Hibana Deaths', 'Hibana EntryKills', 'Hibana EntryDeaths', 'Hibana KOST', 'Hibana HS', 'Hibana MKills', 'Hibana Trades', 'Hibana Clutch', 'Hibana Plants', 'Hibana Rounds', 'Jackal Rating', 'Jackal Kills', 'Jackal Deaths', 'Jackal EntryKills', 'Jackal EntryDeaths', 'Jackal KOST', 'Jackal HS', 'Jackal MKills', 'Jackal Trades', 'Jackal Clutch', 'Jackal Plants', 'Jackal Rounds', 'Ying Rating', 'Ying Kills', 'Ying Deaths', 'Ying EntryKills', 'Ying EntryDeaths', 'Ying KOST', 'Ying HS', 'Ying MKills', 'Ying Trades', 'Ying Clutch', 'Ying Plants', 'Ying Rounds', 'Zofia Rating', 'Zofia Kills', 'Zofia Deaths', 'Zofia EntryKills', 'Zofia EntryDeaths', 'Zofia KOST', 'Zofia HS', 'Zofia MKills', 'Zofia Trades', 'Zofia Clutch', 'Zofia Plants', 'Zofia Rounds', 'Dokkaebi Rating', 'Dokkaebi Kills', 'Dokkaebi Deaths', 'Dokkaebi EntryKills', 'Dokkaebi EntryDeaths', 'Dokkaebi KOST', 'Dokkaebi HS', 'Dokkaebi MKills', 'Dokkaebi Trades', 'Dokkaebi Clutch', 'Dokkaebi Plants', 'Dokkaebi Rounds', 'Finka Rating', 'Finka Kills', 'Finka Deaths', 'Finka EntryKills', 'Finka EntryDeaths', 'Finka KOST', 'Finka HS', 'Finka MKills', 'Finka Trades', 'Finka Clutch', 'Finka Plants', 'Finka Rounds', 'Lion Rating', 'Lion Kills', 'Lion Deaths', 'Lion EntryKills', 'Lion EntryDeaths', 'Lion KOST', 'Lion HS', 'Lion MKills', 'Lion Trades', 'Lion Clutch', 'Lion Plants', 'Lion Rounds', 'Maverick Rating', 'Maverick Kills', 'Maverick Deaths', 'Maverick EntryKills', 'Maverick EntryDeaths', 'Maverick KOST', 'Maverick HS', 'Maverick MKills', 'Maverick Trades', 'Maverick Clutch', 'Maverick Plants', 'Maverick Rounds', 'Nomad Rating', 'Nomad Kills', 'Nomad Deaths', 'Nomad EntryKills', 'Nomad EntryDeaths', 'Nomad KOST', 'Nomad HS', 'Nomad MKills', 'Nomad Trades', 'Nomad Clutch', 'Nomad Plants', 'Nomad Rounds', 'Gridlock Rating', 'Gridlock Kills', 'Gridlock Deaths', 'Gridlock EntryKills', 'Gridlock EntryDeaths', 'Gridlock KOST', 'Gridlock HS', 'Gridlock MKills', 'Gridlock Trades', 'Gridlock Clutch', 'Gridlock Plants', 'Gridlock Rounds', 'Nokk Rating', 'Nokk Kills', 'Nokk Deaths', 'Nokk EntryKills', 'Nokk EntryDeaths', 'Nokk KOST', 'Nokk HS', 'Nokk MKills', 'Nokk Trades', 'Nokk Clutch', 'Nokk Plants', 'Nokk Rounds', 'Amaru Rating', 'Amaru Kills', 'Amaru Deaths', 'Amaru EntryKills', 'Amaru EntryDeaths', 'Amaru KOST', 'Amaru HS', 'Amaru MKills', 'Amaru Trades', 'Amaru Clutch', 'Amaru Plants', 'Amaru Rounds', 'Kali Rating', 'Kali Kills', 'Kali Deaths', 'Kali EntryKills', 'Kali EntryDeaths', 'Kali KOST', 'Kali HS', 'Kali MKills', 'Kali Trades', 'Kali Clutch', 'Kali Plants', 'Kali Rounds', 'Iana Rating', 'Iana Kills', 'Iana Deaths', 'Iana EntryKills', 'Iana EntryDeaths', 'Iana KOST', 'Iana HS', 'Iana MKills', 'Iana Trades', 'Iana Clutch', 'Iana Plants', 'Iana Rounds', 'Ace Rating', 'Ace Kills', 'Ace Deaths', 'Ace EntryKills', 'Ace EntryDeaths', 'Ace KOST', 'Ace HS', 'Ace MKills', 'Ace Trades', 'Ace Clutch', 'Ace Plants', 'Ace Rounds', 'Zero Rating', 'Zero Kills', 'Zero Deaths', 'Zero EntryKills', 'Zero EntryDeaths', 'Zero KOST', 'Zero HS', 'Zero MKills', 'Zero Trades', 'Zero Clutch', 'Zero Plants', 'Zero Rounds', 'Flores Rating', 'Flores Kills', 'Flores Deaths', 'Flores EntryKills', 'Flores EntryDeaths', 'Flores KOST', 'Flores HS', 'Flores MKills', 'Flores Trades', 'Flores Clutch', 'Flores Plants', 'Flores Rounds', 'Osa Rating', 'Osa Kills', 'Osa Deaths', 'Osa EntryKills', 'Osa EntryDeaths', 'Osa KOST', 'Osa HS', 'Osa MKills', 'Osa Trades', 'Osa Clutch', 'Osa Plants', 'Osa Rounds', 'Sens Rating', 'Sens Kills', 'Sens Deaths', 'Sens EntryKills', 'Sens EntryDeaths', 'Sens KOST', 'Sens HS', 'Sens MKills', 'Sens Trades', 'Sens Clutch', 'Sens Plants', 'Sens Rounds', 'Grim Rating', 'Grim Kills', 'Grim Deaths', 'Grim EntryKills', 'Grim EntryDeaths', 'Grim KOST', 'Grim HS', 'Grim MKills', 'Grim Trades', 'Grim Clutch', 'Grim Plants', 'Grim Rounds', 'Brava Rating', 'Brava Kills', 'Brava Deaths', 'Brava EntryKills', 'Brava EntryDeaths', 'Brava KOST', 'Brava HS', 'Brava MKills', 'Brava Trades', 'Brava Clutch', 'Brava Plants', 'Brava Rounds']
# Select the desired sheet
sheetName = 'Stats'
opSheetName = 'Operator Stats'
excelMainSheet = workbook[sheetName]
excelOpStatSheet = workbook[opSheetName]
#Define variables
excelUsername = []
excelRating = []
excelKills = []
excelDeaths = []
excelKD = []
excelEntryKill = []
excelEntryDeath = []
excelEntryPlusMinus = []
excelKOST = []
excelKPR = []
excelSRV = []
excelMKills = []
excelTrade = []
excelClutch = []
excelPlants = []
excelHS = []
excelAtk = []
excelDef = []
excelRound = []
excelColUsername = 'A'
excelColRating = 'B'
excelColKills = 'C' 
excelColDeaths = 'D'
excelColKD = 'E'
excelColEK = 'F'
excelColED = 'G'
excelColEntry = 'H'
excelColKOST = 'I'
excelColKPR = 'J'
excelColSRV = 'K'
excelColMKills = 'L'
excelColTrade = 'M'
excelColClutch = 'N'
excelColPlants = 'O'
excelColHS = 'P'
excelColRound = 'Q'

#Array of the column names to make it easier to access individual columns below
excelCols = [excelColUsername,excelColRating, excelColKills,excelColDeaths,excelColKD,excelColEK,excelColED,excelColEntry,excelColKOST,excelColKPR,excelColSRV,excelColMKills,excelColTrade,excelColClutch,excelColPlants,excelColHS,excelColRound]

#For each column, save all of data to an array
#Starts at the second cell because the Header is written in the Excel File
for cell in excelMainSheet[excelCols[0]]:
    excelUsername.append(cell.value)
excelUsername = excelUsername[1:]
for cell in excelMainSheet[excelCols[1]]:
    excelRating.append(cell.value)
excelRating = excelRating[1:]
for cell in excelMainSheet[excelCols[2]]:
    excelKills.append(cell.value)
excelKills = excelKills[1:]
for cell in excelMainSheet[excelCols[3]]:
    excelDeaths.append(cell.value)
excelDeaths = excelDeaths[1:]
for cell in excelMainSheet[excelCols[4]]:
    excelKD.append(cell.value)
excelKD = excelKD[1:]
for cell in excelMainSheet[excelCols[5]]:
    excelEntryKill.append(cell.value)
excelEntryKill = excelEntryKill[1:]
for cell in excelMainSheet[excelCols[6]]:
    excelEntryDeath.append(cell.value)
excelEntryDeath = excelEntryDeath[1:]
for cell in excelMainSheet[excelCols[7]]:
    excelEntryPlusMinus.append(cell.value)
excelEntryPlusMinus = excelEntryPlusMinus[1:]
for cell in excelMainSheet[excelCols[8]]:
    excelKOST.append(cell.value)
excelKOST = excelKOST[1:]
for cell in excelMainSheet[excelCols[9]]:
    excelKPR.append(cell.value)
excelKPR = excelKPR[1:]
for cell in excelMainSheet[excelCols[10]]:
    excelSRV.append(cell.value)
excelSRV = excelSRV[1:]
for cell in excelMainSheet[excelCols[11]]:
    excelMKills.append(cell.value)
excelMKills = excelMKills[1:]
for cell in excelMainSheet[excelCols[12]]:
    excelTrade.append(cell.value)
excelTrade = excelTrade[1:]
for cell in excelMainSheet[excelCols[13]]:
    excelClutch.append(cell.value)
excelClutch = excelClutch[1:]
for cell in excelMainSheet[excelCols[14]]:
    excelPlants.append(cell.value)
excelPlants = excelPlants[1:]
for cell in excelMainSheet[excelCols[15]]:
    excelHS.append(cell.value)
excelHS = excelHS[1:]
for cell in excelMainSheet[excelCols[16]]:
    excelRound.append(cell.value)
excelRound = excelRound[1:]


#Takes in the length of the ExcelUsername for 'Operator Stats' to read that many lines from the Operator Stats and read down the row
#Creates a 2d array of width of array of ExcelUsername and Length of # of players
opArray = []

if len(excelUsername) != 0:
    for i in range(len(excelUsername)):
        tempArr = []
        for cell in excelOpStatSheet[2+i]:
            tempArr.append(cell.value)
        opArray.append(tempArr)












#Definiton of rating system, same as used in basic stats, !!SHOULD BECOME UNIFORM!!
def ratingSys(rKills, rKD, rMK, rEntry, rPlants, rClutch, rKOST, rSRV, rRounds):
    rating = (rKD**2 + 0.4*(rKills) + 0.15*rMK)/rRounds + 0.75*(rEntry)/rRounds + (rPlants + rClutch)/rRounds + rKOST + rSRV/3
    return rating
    
#Creates Structs for each player and their stats
class basicStats:
    #Initializes variables, (Why do I have to do this?)
    def __init__(self, username, kills, deaths, kD, eKills, eDeaths, entry, kOST, kPR, sRV, mKills, trade, clutch, plants, defuse, hsPercent, favAtk, favDef, rounds, opK, opD, opKD, opEK, opED, opEntry, opKOST, opKPR, opSRV, opMKills, opTrade, opClutch, opPlants, opDefuse, opHS, opRounds, mapPlayed, team1Score, team2Score, team1Name, team2Name):
        self.username = username
        self.kills = kills
        self.deaths = deaths
        self.kD = kD
        self.eKills = eKills
        self.eDeaths = eDeaths
        self.entry = entry
        self.kOST = kOST
        self.kPR = kPR
        self.sRV = sRV
        self.mKills = mKills
        self.trade = trade
        self.clutch = clutch
        self.plants = plants
        self.defuse = defuse
        self.hsPercent = hsPercent
        self.favAtk = favAtk
        self.favDef = favDef
        self.rounds = rounds
        self.opK = opK
        self.opD = opD
        self.opKD = opKD
        self.opEK = opEK
        self.opED = opED
        self.opEntry = opEntry
        self.opKOST = opKOST
        self.opKPR = opKPR
        self.opSRV = opSRV
        self.opMKills = opMKills
        self.opTrade = opTrade
        self.opClutch = opClutch
        self.opPlants = opPlants
        self.opDefuse = opDefuse
        self.opHS = opHS
        self.opRounds = opRounds
        self.mapPlayed = mapPlayed
        self.team1Score = team1Score
        self.team2Score = team2Score
        self.team1Name = team1Name
        self.team2Name = team2Name
        
#Rating System, very basic
#Same rating system as before
    # def rating(self):
    #     rating = (self.KD**2 + 0.4*(self.Kills) + 0.15*self.MKills)/self.Rounds + 0.75*(self.Entry)/self.Rounds + (self.Plants + self.Clutch)/self.Rounds + self.KOST + self.SRV/3
    #     return rating
#Prints out all 'relevant' stats, in similar format as siegeGG, adds multikills and trades for greater visibility
    def printIndivStat(self, intro):
        #defines the rating in this def
        rating = ratingSys(self.kills, self.kD, self.mKills, self.entry, self.plants, self.clutch, self.kOST, self.sRV, self.rounds)
        #if the first user printed, prints a header
        
        #Top part of the print, gives the map, and score and header
        if(intro == 1):
            print('Map: ' + self.mapPlayed)
            print('')
            print(self.team1Name + ' - ' + self.team2Name)
            print(str(self.team1Score) + '-' + str(self.team2Score))
            formattedString = "{:<15} | {:<6} | {:<10} | {:<8} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<10} | {:<10}".format(
                'Username', 'Rating', 'K-D(KD)', 'Entry', 'KOST', 'KPR', 'SRV', '1vX', 'Plants', 'Multis', 'Trades', 'HS', 'Attacker', 'Defender', '', ''
            )
            #copies the underlined text to the length of the text above and creates a line
            underlinedString = formattedString + '\n' + '-' * len(formattedString)
            print(underlinedString)

        #formats the plus or minus in front of the KD and entry to make it + or -
        plusMinus = self.kills - self.deaths
        if(plusMinus > 0):
            strKD = str(self.kills) + '-' + str(self.deaths) + '(+' + str(plusMinus)+')'
        else:
            strKD = str(self.kills) + '-' + str(self.deaths) + '(' + str(plusMinus)+')'
        
        #Same formatting for Entry Stats
        ePlusMinus = self.eKills - self.eDeaths
        if(ePlusMinus > 0):
            strEntry = str(int(self.eKills)) + '-' + str(int(self.eDeaths)) + '(+'+str(int(ePlusMinus))+')'
        else:
            strEntry = str(int(self.eKills)) + '-' + str(int(self.eDeaths)) + '('+str(int(ePlusMinus))+')'
        #formatting the text for each user and prints
        formatKOST = "{:.2f}".format(self.kOST)
        formatKPR = "{:.2f}".format(self.kPR)
        formatRating = "{:.2f}".format(rating)
        formattedString = "{:<15} | {:<6} | {:<10} | {:<8} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<10} | {:<10}".format(
            self.username,
            formatRating,
            strKD,
            strEntry,
            formatKOST,
            formatKPR,
            str(int(self.sRV*100)) + '%',
            int(self.clutch),
            int(self.plants),
            int(self.mKills),
            int(self.trade),
            str(int(self.hsPercent)) + '%',
            self.favAtk,
            self.favDef
        )
        print(formattedString)
    #Define ratings for Operator Ratings for a player
    def operatorRating(self):
        operatorRating = np.zeros(opNumbers)
        #rating = (1.5*self.KD + 0.25*(self.Kills) + 0.15*self.MKills)/self.Rounds + 0.75*(self.Entry)/self.Rounds + (self.Plants + self.Clutch)/self.Rounds + self.KOST + self.SRV/3
        for j in (range(opNumbers)):
            if self.opRounds[j] == 0:
                operatorRating[j] = 0
            else:
                operatorRating[j] = ratingSys(self.opK[j], self.opKD[j], self.opMKills[j], self.opEntry[j], self.opPlants[j], self.opClutch[j], self.opKOST[j], self.opSRV[j], self.opRounds[j])
        return operatorRating
    #Simple Way to read all player Operator Rating, just in python currently, but could be phased out
    def allOps(self):
        Op = self.operatorRating()
        for k in (range(opNumbers)):
            number_str = Op[k]
            roundedRating = "{:.2f}".format(float(number_str))
            print(numToOperators[k] + ': ' + roundedRating)
    #Similar to above, can look at individual player and full rating for a single operator, python only, either phase out or can be used in maybe a different aspect
    #Very similar formatting to the full list for a single map, should be merged
    def singleOperatorStats(self, inputStr):
        operatorValue = operatorsToNum[inputStr]
        print('\n')
        opsRate = self.operatorRating()
        
        formattedString = "{:<12} | {:<10} | {:<6} | {:<10} | {:<8} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6}".format(
                'Player','Op Name', 'Rating', 'K-D(KD)', 'Entry', 'KOST', 'KPR', 'SRV', '1vX', 'Plants', 'Multis', 'Trades', 'HS', 'Rounds', '')
        underlinedString = formattedString + '\n' + '-' * len(formattedString)
        print(underlinedString)
        plusMinus = self.opK[operatorValue] - self.opD[operatorValue]
        if(plusMinus > 0):
            strKD = str(int(self.opK[operatorValue])) + '-' + str(int(self.opD[operatorValue])) + '(+' + str(int(plusMinus))+')'
        else:
            strKD = str(int(self.opK[operatorValue])) + '-' + str(int(self.opD[operatorValue])) + '(' + str(int(plusMinus))+')'
        
        ePlusMinus = self.opEK[operatorValue] - self.opED[operatorValue]
        if(ePlusMinus > 0):
            strEntry = str(int(self.opEK[operatorValue])) + '-' + str(int(self.opED[operatorValue])) + '(+'+str(int(ePlusMinus))+')'
        else:
            strEntry = str(int(self.opEK[operatorValue])) + '-' + str(int(self.opED[operatorValue])) + '('+str(int(ePlusMinus))+')'
        #formatting the text for each user and prints
        if self.opK[operatorValue] == 0:
            hs = 0
        else:
            hs = self.opHS[operatorValue]/self.opK[operatorValue]*100
        formatKOST = "{:.2f}".format(self.opKOST[operatorValue])
        formatKPR = "{:.2f}".format(self.opKPR[operatorValue])
        formatRating = "{:.2f}".format(opsRate[operatorValue])
        
        formattedString = "{:<12} | {:<10} | {:<6} | {:<10} | {:<8} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6}".format(
            self.username,
            inputStr,
            formatRating,
            strKD,
            strEntry,
            formatKOST,
            formatKPR,
            str(int(self.opSRV[operatorValue]*100)) + '%',
            int(self.opClutch[operatorValue]),
            int(self.opPlants[operatorValue]),
            int(self.opMKills[operatorValue]),
            int(self.opTrade[operatorValue]),
            str(int(hs)) + '%',
            str(int(self.opRounds[operatorValue]))
        )
        print(formattedString)
            



#Parses json file for all stats
#Unable to determine how the disable works !!NEEDS TESTING!!
def singleMap(dict):
    #define all variables to pass through/use
    team1Score = 0
    team2Score = 0
    #Takes the team names from the first round
    team1Name = dict["rounds"][0]["teams"][0]["name"]
    team2Name = dict["rounds"][0]["teams"][1]["name"]
    usernameList = []
    usernameLookup = {}
    killAmount = []
    deathAmount = []
    hsPercent = []
    roundCount = []
    atkMain = []
    defMain = []
    #Note: The operator arrays are filled to -1, this is because if it was initialized at 0, could show a false positive of mute being played because he is operator 0
    roundOps = np.full(10,-1)
    dTotalOps = np.array([])
    aTotalOps = np.array([])
    aRoundOps = np.full(10,-1)
    dRoundOps = np.full(10,-1)
    aMainOp = np.array([])
    dMainOp = np.array([])
    winningTeamMembers = []
    entryKills = np.zeros(10)
    entryDeaths = np.zeros(10)
    plants = np.zeros(10)
    defusal = np.zeros(10)
    kOSTRounds = np.zeros(10)
    kOSTTotal=np.zeros(10)
    kOSTSurv = np.zeros(10)
    clutches = np.zeros(10)
    multikills = np.zeros(10)
    trades = np.zeros(10)
    opKills = np.zeros((10,68))
    opDeaths = np.zeros((10,68))
    opEKills = np.zeros((10,68))
    opEDeaths = np.zeros((10,68))
    opHS = np.zeros((10,68))
    opPlants = np.zeros((10,68))
    opDefusal = np.zeros((10,68))
    opKOST = np.zeros((10,68))
    opKOSTRound = np.zeros((10,68))
    opClutches = np.zeros((10,68))
    opMultikills = np.zeros((10,68))
    opTrades = np.zeros((10,68))
    opRounds = np.zeros((10,68))
    
    #Takes the first round map name to output
    map = dict["rounds"][0]["map"]["name"]
    #takes the total number of kills in the map and hs percent and rounds played
    for i in range(10):
        #Appends list at the end of rounds to track the basic stats that is given, could be improved if needed
        usernameList.append(dict["stats"][i]["username"])
        killAmount.append(dict["stats"][i]["kills"])
        deathAmount.append(dict["stats"][i]["deaths"])
        hsPercent.append(dict["stats"][i]["headshotPercentage"])
        roundCount.append(dict["stats"][i]["rounds"])
        #creates a lookup table for each player, 0-9 based on name and how incrememented in the json file
        #Makes it easier to upload names to further stats, because can organize players
        usernameLookup[usernameList[i]] = i
    #large loop to look at all rounds
    #!! Restructing Note, should do a double for loop where outer loop is rounds, and inner loops in actions in rounds !!
    for i in range(len(dict["rounds"])):
        actions = len(dict["rounds"][i]["matchFeedback"])
        
        #actions is number of occurences in the 'main phase' of each round, contains everything that a player can do to impact each round
        #!!2 Different logics between while and for loop, should restructure!!
        
        #Updates team score each round, so the final score is output at the end of all rounds
        team1Score = dict["rounds"][i]["teams"][0]["score"]
        team2Score = dict["rounds"][i]["teams"][1]["score"]
        actions = len(dict["rounds"][i]["matchFeedback"])
        #loops through all the actions looking for the first kill to occur
        aRoundOps = np.full(10,-1)
        dRoundOps = np.full(10,-1)
        if(dict["rounds"][i]["teams"][0]["won"] == True):
            WinningTeam = 0
        else:
            WinningTeam = 1
        winningTeamMembers = []
        clutchPlayer = ''
        clutchAlive = 5
        for v in range(len(dict["rounds"][i]["players"])):
            #tracks clutching based on if your team won and you were the only player on your team alive
            
            if(dict["rounds"][i]["players"][v]["teamIndex"] == WinningTeam):
                winningTeamMembers.append(dict["rounds"][i]["players"][v]["username"])
            #start with 5 players alive on the winning team, if a winning team member dies, then reduce the number alive
            #add winning round members
            
            
            #if the operator that is selected by the player is > 33 it is an attacker by my dictionary, otherwise its a defender
            #if an attacker, needs to check if a repick occurs
            if(operatorsToNum[dict["rounds"][i]["players"][v]["operator"]["name"]] > 33):
                #Puts the Player on the operator that they played that round in their 'spot' in the array, and given a numerical value from the dict 'Operators'
                roundOps[usernameLookup[dict["rounds"][i]["players"][v]["username"]]] = operatorsToNum[dict["rounds"][i]["players"][v]["operator"]["name"]]
                aRoundOps[usernameLookup[dict["rounds"][i]["players"][v]["username"]]] = operatorsToNum[dict["rounds"][i]["players"][v]["operator"]["name"]]
                #as attackers can swap in prep phase, look for the match feedback for an operator swap and update the value
                #Checks through all the round actions to see if an operator is swapped off
                for c in range(actions):
                    if(dict["rounds"][i]["matchFeedback"][c]["type"]["name"] == "OperatorSwap"):
                        #Updates the value in similar value
                        roundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][c]["username"]]] = operatorsToNum[dict["rounds"][i]["matchFeedback"][c]["operator"]["name"]]
                        aRoundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][c]["username"]]] = operatorsToNum[dict["rounds"][i]["matchFeedback"][c]["operator"]["name"]]
                        
            else:
                #Defenders do not change, so the first time seen can be set
                roundOps[usernameLookup[dict["rounds"][i]["players"][v]["username"]]] = operatorsToNum[dict["rounds"][i]["players"][v]["operator"]["name"]]
                dRoundOps[usernameLookup[dict["rounds"][i]["players"][v]["username"]]] = operatorsToNum[dict["rounds"][i]["players"][v]["operator"]["name"]]
            #Hard to read
            #Puts all the operator plays into the 2d array for the Operator Stats Page, Puts the username lookup and and operator location and increments the play count
            
            opRounds[usernameLookup[dict["rounds"][i]["players"][v]["username"]],roundOps[usernameLookup[dict["rounds"][i]["players"][v]["username"]]]] = opRounds[usernameLookup[dict["rounds"][i]["players"][v]["username"]],roundOps[usernameLookup[dict["rounds"][i]["players"][v]["username"]]]] + 1
        #use the dict defined numerical value of the op to add to a total array
        aTotalOps = np.append(aTotalOps, aRoundOps)
        dTotalOps = np.append(dTotalOps, dRoundOps)
        #Loops for opening kill
        firstKill = 0
        kOSTRounds = np.zeros(10)
        #If did nothing but lived, should have 1
        kOSTSurv = np.ones(10)
        opKOSTRound = np.zeros((10,opNumbers))
        for j in range(actions):
            
            if firstKill == 0 and dict["rounds"][i]["matchFeedback"][j]["type"]["name"] == "Kill":
                firstKill = firstKill + 1
                entryKills[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] = entryKills[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] + 1
                entryDeaths[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]] = entryDeaths[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]] + 1
                opEKills[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]],roundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]]] = opEKills[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]],roundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]]] + 1
                opEDeaths[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]],roundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]]] = opEDeaths[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]],roundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]]] + 1
            #gives the planter a point if they get defuser down and what operator they played
            if dict["rounds"][i]["matchFeedback"][j]["type"]["name"] == "DefuserPlantComplete":
                plants[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] = plants[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] + 1
                opPlants[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]],roundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]]] = opPlants[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]],roundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]]]+1
            if(dict["rounds"][i]["matchFeedback"][j]["type"]["name"] == "Kill"):
                #If a kill happens, the round is counted for KOST, but not survival
                kOSTRounds[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] = 1
                kOSTSurv[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]] = 0
                #looks for trades, when a user dies, capture tOD and who killed them
                #then iterate through until the time is past the time to trade to see if the trade occurs
                timeOfDeath = dict["rounds"][i]["matchFeedback"][j]["timeInSeconds"]
                userToBeTraded = dict["rounds"][i]["matchFeedback"][j]["username"]
                l=0
                #If the time happens, loops through all actions to see if it is within the 10 second trade time, otherwise go to next action
                while(timeOfDeath-timeToTrade < dict["rounds"][i]["matchFeedback"][j+l]["timeInSeconds"]):
                    if(dict["rounds"][i]["matchFeedback"][j+l]["type"]["name"] == "Kill" and dict["rounds"][i]["matchFeedback"][j+l]["target"] == userToBeTraded):
                        kOSTRounds[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]] = 1
                        trades[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]] = trades[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]] + 1
                        opTrades[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]],roundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]]] = opTrades[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]],roundOps[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]]] + 1
                        break
                    else:
                        l = l+1
                        if(j+l == actions):
                            break
                
            #adds plant if get plant down
            elif(dict["rounds"][i]["matchFeedback"][j]["type"]["name"] == "DefuserPlantComplete"):
                kOSTRounds[usernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] = 1
        
        
            

        #tracks KOST (Kills, Objectives[plants], Survival, Trades)
        
        #looks for survival rate and trades
        
            
        #adds a round where you add to your kost, makes sure not to duplicate if 2 things occur in a round
        #Adds rounds to if they survived
        kOSTRounds = kOSTRounds + kOSTSurv
        
        #If higher than one, sets back to 1 to not overrate the KOST rounds if someone does more than one KOST action in a round
        for n in range(len(kOSTRounds)):
            if(kOSTRounds[n] > 1):
                kOSTRounds[n] = 1
            opKOSTRound[n, roundOps[n]] = kOSTRounds[n]
        #Adds the KOST to total number of rounds that players achieved during the number of rounds
        kOSTTotal = kOSTTotal + kOSTRounds
        opKOST = opKOST + opKOSTRound

        #tracks each user for their multikills over a game
        #Checks for HS percentage and Kills for operators per round, also looks to see if number of kills was greater than 1
        for g in range(len(dict["rounds"][i]["stats"])):
            
            opKills[usernameLookup[dict["rounds"][i]["stats"][g]["username"]],roundOps[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] = opKills[usernameLookup[dict["rounds"][i]["stats"][g]["username"]],roundOps[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] + dict["rounds"][i]["stats"][g]["kills"]
            opHS[usernameLookup[dict["rounds"][i]["stats"][g]["username"]],roundOps[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] = opHS[usernameLookup[dict["rounds"][i]["stats"][g]["username"]],roundOps[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] + dict["rounds"][i]["stats"][g]["headshots"]
            if(dict["rounds"][i]["stats"][g]["died"] == True):
                opDeaths[usernameLookup[dict["rounds"][i]["stats"][g]["username"]],roundOps[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] = opDeaths[usernameLookup[dict["rounds"][i]["stats"][g]["username"]],roundOps[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] + 1
            if(dict["rounds"][i]["stats"][g]["kills"] > 1):
                multikills[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]] = multikills[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]] + 1
                opMultikills[usernameLookup[dict["rounds"][i]["stats"][g]["username"]], roundOps[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] = opMultikills[usernameLookup[dict["rounds"][i]["stats"][g]["username"]], roundOps[usernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] + 1
        
        
            
        
        #if the winning team member died, then reduce the number, else assign that person the clutch player, if nobody is alive on a team then this doesnt get added to clutch total
        for a in range(len(winningTeamMembers)):
            for b in range(len(dict["rounds"][i]["stats"])):
                if(winningTeamMembers[a] == dict["rounds"][i]["stats"][b]["username"] and dict["rounds"][i]["stats"][b]["died"] == True):
                    clutchAlive = clutchAlive - 1
                else:
                    clutchPlayer = dict["rounds"][i]["stats"][b]["username"]
        #Check again to see that only 1 player was alive on the winning team to confirm that it is a clutch and give it to player and the operator they played that round
        if(clutchAlive == 1):
            clutches[usernameLookup[clutchPlayer]] = clutches[usernameLookup[clutchPlayer]] + 1
            opClutches[usernameLookup[clutchPlayer],roundOps[usernameLookup[clutchPlayer]]] = opClutches[usernameLookup[clutchPlayer],roundOps[usernameLookup[clutchPlayer]]] + 1

        
        
        #find all main ops by using array math to loop through and find each players most played operator
    
    for z in range(len(dict["rounds"][i]["players"])):
        aMainOp = np.array([])
        dMainOp = np.array([])
        #Separate players into their individual operators to check for main operator that a single person
        for q in range(len(dict["rounds"])):
            aMainOp = np.append(aMainOp, aTotalOps[q*10 + z])
            aMainOp = aMainOp[aMainOp>=0]
            dMainOp = np.append(dMainOp, dTotalOps[q*10 + z])
            dMainOp = dMainOp[dMainOp>=0]
        #add each users main op back to main array
        #Create an array to check what the main operator each person played
        aMain = statistics.mode(aMainOp)
        aMain = int(aMain)
        dMain = statistics.mode(dMainOp)
        atkMain.append(numToOperators[aMain])
        defMain.append(numToOperators[dMain])
    
        
    #all return values are array values of # of players
    return [usernameList,killAmount,deathAmount,entryKills,entryDeaths,kOSTTotal, hsPercent,multikills,trades, clutches, plants,defusal,atkMain,defMain,roundCount, opKills, opDeaths, opEKills, opEDeaths, opKOST, opHS, opMultikills, opTrades, opClutches, opPlants,opDefusal, opRounds, map, team1Score, team2Score, team1Name, team2Name]
        


#open the json file i am parsing
with open("C:\\Users\\jakeg\\OneDrive\\Desktop\\r6-dissect-v0.11.1-windows-amd64\\scrim6.json", 'r') as f:
    my_dict = json.load(f)

#function to parse data
[users, kills, deaths, eKills, eDeaths, kOST, hs, mk, trade, clutch, plant, defuse, attackerMain, defenderMain, rounds, operatorKills, operatorDeaths, operatorEntryKills, operatorEntryDeaths, operatorKOST, operatorHS, operatorMKills, operatorTrades, operatorClutch, operatorPlants, operatorDefusal, operatorRounds, mapPlayed, team1Score, team2Score, team1Name, team2Name] = singleMap(my_dict)

#use for loop as basic tool to print all player data, similar to siege GG
players = []

#Format output of function to what is necessary to the Basic Stats class
for i in range(len(users)):
    #Intermediate values to make a potentially undefined value, or additional info necessary, like entry being subtracted
    #IE, if a person played 1 round, got 2 kills and didnt die, KD would be infinite if not reset
    operatorEntry = np.zeros(opNumbers)
    operatorKD = np.zeros(opNumbers)
    operatorKOSTAmount = np.zeros(opNumbers)
    operatorKPR = np.zeros(opNumbers)
    operatorSRV = np.zeros(opNumbers)
    kd = kills[i]/deaths[i]
    entry = eKills[i]-eDeaths[i]
    kOSTAmount = kOST[i]/rounds[i]
    kPR = kills[i]/rounds[i]
    sRV = (rounds[i]-deaths[i])/rounds[i]
    #Repeats for number of operators as similar stats
    for j in range(opNumbers):
        operatorEntry[j] = operatorEntryKills[i][j] - operatorEntryDeaths[i][j]
        if operatorDeaths[i][j] == 0:
            operatorKD[j] = operatorKills[i][j]
        else:
            operatorKills[i][j]/operatorDeaths[i][j]
        if operatorRounds[i][j] == 0:
            operatorKOSTAmount[j] = 0
            operatorKPR[j] = 0
            operatorSRV[j] = 0
        else:
            operatorKOSTAmount[j] = operatorKOST[i][j]/operatorRounds[i][j]
            operatorKPR[j] = operatorKills[i][j]/operatorRounds[i][j]
            operatorSRV[j] = (operatorRounds[i][j] - operatorDeaths[i][j])/operatorRounds[i][j]

    player = basicStats(users[i],kills[i],deaths[i],kd,eKills[i],eDeaths[i],entry,kOSTAmount,kPR,sRV,mk[i],trade[i],clutch[i],plant[i],defuse[i],hs[i],attackerMain[i], defenderMain[i], rounds[i],operatorKills[i], operatorDeaths[i],operatorKD, operatorEntryKills[i], operatorEntryDeaths[i], operatorEntry, operatorKOSTAmount, operatorKPR, operatorSRV, operatorMKills[i], operatorTrades[i], operatorClutch[i], operatorPlants[i], operatorDefusal[i], operatorHS[i], operatorRounds[i], mapPlayed, team1Score, team2Score, team1Name, team2Name)
    players.append(player)
    #Prints to terminal
    if(i==0):
        start = 1
    else:
        start = 0
    players[i].printIndivStat(start)
#Can be an outdated value
excelUserLists = len(excelUsername)
#For the number of users in the current match
for i in range(len(users)):
    matching = 0
    matchingValue = 0
    #Checks these names against the list in Excel
    for j in range(len(excelUsername)):
        #If it matches, can confirm that they have played prior
        if users[i] == opArray[j][0]:
            matchingValue = j
            matching = matching + 1
    #If doesnt match, add the player to the list
    if matching == 0:
        array = []
        operatorRating = []
        singleUser = []
        tempKD = 0
        #Same stopping of infinite values and replacing or undefined values
        for b in range(len(operatorsToNum)):
            if operatorKills[i][b] == 0:
                tempHS = 0
            else:
                tempHS = operatorHS[i][b]/operatorKills[i][b]
            if operatorDeaths[i][b] == 0:
                tempKD = operatorKills[i][b]
            else:
                tempKD = operatorKills[i][b]/operatorDeaths[i][b]
            if operatorRounds[i][b] == 0:
                tempKOST = 0
                tempSRV = 0
                tempRating = 0
            else:
                tempKOST = operatorKOST[i][b]/operatorRounds[i][b]
                tempSRV = (operatorRounds[i][b] - operatorDeaths[i][b])/operatorRounds[i][b]
                tempRating = ratingSys(operatorKills[i][b],tempKD, operatorMKills[i][b], operatorEntryKills[i][b] - operatorEntryDeaths[i][b], operatorPlants[i][b], operatorClutch[i][b], tempKOST, tempSRV, operatorRounds[i][b] )

            #Array to add to excel when all compiled
            array.append(round((tempRating),2))
            operatorRating.append(round((tempRating),2))
            array.append(operatorKills[i][b])
            array.append(operatorDeaths[i][b])
            array.append(operatorEntryKills[i][b])
            array.append(operatorEntryDeaths[i][b])
            array.append(round(tempKOST,2))
            array.append(round(tempHS,2))
            array.append(operatorMKills[i][b])
            array.append(operatorTrades[i][b])
            array.append(operatorClutch[i][b])
            array.append(operatorPlants[i][b])
            array.append(operatorRounds[i][b])
        singleUser.append(users[i])
        #Within the Operators, looking at rounds for the operator plays and split between attackers and defenders
        operatorDefRounds = operatorRounds[i][0:33]
        operatorDefRating = operatorRating[0:33]
        operatorAtkRounds = operatorRounds[i][34:]
        operatorAtkRating = operatorRating[34:]
        
        #Find the max of the Rounds and ratings, these arrays are of slightly different types, but calculate the same thing
        maxDefRounds = np.argmax(operatorDefRounds)
        maxAtkRounds = np.argmax(operatorAtkRounds)
        maxDefRating = max(operatorDefRating)
        maxAtkRating = max(operatorAtkRating)
        
        #Append these values to the user array for excel
        singleUser.append(numToOperators[34+maxAtkRounds])
        singleUser.append(operatorRounds[i][34+maxAtkRounds])
        singleUser.append(numToOperators[maxDefRounds])
        singleUser.append(operatorRounds[i][maxDefRounds])
        
        singleUser.append(numToOperators[34+operatorAtkRating.index(maxAtkRating)])
        singleUser.append(operatorRating[34+operatorAtkRating.index(maxAtkRating)])
        singleUser.append(numToOperators[operatorDefRating.index(maxDefRating)])
        singleUser.append(operatorRating[operatorDefRating.index(maxDefRating)])

        #Adds the operator values to the single User arrays
        for f in range(len(array)):
            singleUser.append(array[f])
        opArray.append(singleUser)
    #This is the values when a player already exists in excel
    else:
        tempAtkRating = []
        tempDefRating = []
        tempAtkRounds = []
        tempDefRounds = []
        #Add values the operator array that will be put into the excel
        #Updates the operator values that already exist in the values
        #Offsets based on excel to access the correct data
        for d in range(len(operatorsToNum)):
            #This is checking headshots and headshot value
            if opArray[matchingValue][10+12*d] == 0:
                opArray[matchingValue][15+12*d] = 0
            else:
                opArray[matchingValue][15+12*d] = round((opArray[matchingValue][15+12*d]*opArray[matchingValue][10+12*d] + operatorHS[i][d])/(opArray[matchingValue][10+12*d] + operatorKills[i][d]),2)
            #Updates all values that are empirical
            opArray[matchingValue][10+12*d] = opArray[matchingValue][10+12*d] + operatorKills[i][d]
            opArray[matchingValue][11+12*d] = opArray[matchingValue][11+12*d] + operatorDeaths[i][d]
            opArray[matchingValue][12+12*d] = opArray[matchingValue][12+12*d] + operatorEntryKills[i][d]
            opArray[matchingValue][13+12*d] = opArray[matchingValue][13+12*d] + operatorEntryDeaths[i][d]
            opArray[matchingValue][16+12*d] = opArray[matchingValue][16+12*d] + operatorMKills[i][d]
            opArray[matchingValue][17+12*d] = opArray[matchingValue][17+12*d] + operatorTrades[i][d]
            opArray[matchingValue][18+12*d] = opArray[matchingValue][18+12*d] + operatorClutch[i][d]
            opArray[matchingValue][19+12*d] = opArray[matchingValue][19+12*d] + operatorPlants[i][d]
            opArray[matchingValue][20+12*d] = opArray[matchingValue][20+12*d] + operatorRounds[i][d]
            #Calculates Operator KD
            if opArray[matchingValue][11+12*d] == 0:
                tempKD = opArray[matchingValue][10+12*d]
            else:
                tempKD = opArray[matchingValue][10+12*d]/opArray[matchingValue][11+12*d]
            #Calculates survival
            if opArray[matchingValue][20+12*d] == 0:
                opArray[matchingValue][14+12*d] = 0
                tempSRV = 0
                tempRating = 0
            else:
                opArray[matchingValue][14+12*d] = round((opArray[matchingValue][14+12*d]*opArray[matchingValue][20+12*d] + operatorKOST[i][d])/(opArray[matchingValue][20+12*d] + operatorRounds[i][d]),2)
                tempSRV = (opArray[matchingValue][20+12*d] - opArray[matchingValue][11+12*d])/opArray[matchingValue][20+12*d]
                tempRating = round(ratingSys(opArray[matchingValue][10+12*d], tempKD, opArray[matchingValue][16+12*d], opArray[matchingValue][12+12*d] - opArray[matchingValue][13+12*d], opArray[matchingValue][19+12*d], opArray[matchingValue][18+12*d], opArray[matchingValue][14+12*d], tempSRV, opArray[matchingValue][20+12*d]),2)
            opArray[matchingValue][9+12*d] = tempRating
            #Updates the attacker ratings per operator
            if d >= 34:
                tempAtkRating.append(opArray[matchingValue][9+12*d])
                
                
                tempAtkRounds.append(opArray[matchingValue][20+12*d])
            else:
                
                tempDefRating.append(opArray[matchingValue][9+12*d])
                tempDefRounds.append(opArray[matchingValue][20+12*d])
        #Puts the max ratings and rounds into the array
        maxAtkRating = max(tempAtkRating)
        maxDefRating = max(tempDefRating)
        maxAtkRounds = max(tempAtkRounds)
        maxDefRounds = max(tempDefRounds)
        opArray[matchingValue][1] = numToOperators[34+tempAtkRounds.index(maxAtkRounds)]
        opArray[matchingValue][2] = maxAtkRounds
        opArray[matchingValue][3] = numToOperators[tempDefRounds.index(maxDefRounds)]
        opArray[matchingValue][4] = maxDefRounds
        opArray[matchingValue][5] = numToOperators[34+tempAtkRating.index(maxAtkRating)]
        opArray[matchingValue][6] = maxAtkRating
        opArray[matchingValue][7] = numToOperators[tempDefRating.index(maxDefRating)]
        opArray[matchingValue][8] = maxDefRating
            
            



#Check new or returning players for the Stats Excel Page
for column_index, value in enumerate(operatorHeader, start=1):
    cell = excelOpStatSheet.cell(row=1, column=column_index)
    cell.value = value

for row_index, row in enumerate(opArray, start=2):
    for column_index, value in enumerate(row, start=1):
        cell = excelOpStatSheet.cell(row=row_index, column=column_index)
        cell.value = value
#Checks the users to see if they already match
for i in range(len(users)):
    matching = 0
    matchingValue = -1
    for j in range(len(excelUsername)):
        if excelUsername[j] == users[i]:
            matching = matching + 1
            matchingValue = j
    #If do not match append to the arrays
    if matching == 0:
        excelUsername.append(users[i])
        excelRating.append(round(ratingSys(kills[i], kills[i]/deaths[i], mk[i], eKills[i] - eDeaths[i], plant[i], clutch[i], kOST[i]/rounds[i], (rounds[i]-deaths[i])/rounds[i], rounds[i]),2))
        excelKills.append(kills[i])
        excelDeaths.append(deaths[i])
        excelKD.append(round(kills[i]/deaths[i], 2))
        excelEntryKill.append(eKills[i])
        excelEntryDeath.append(eDeaths[i])
        excelEntryPlusMinus.append(eKills[i] - eDeaths[i])
        excelKOST.append(round(kOST[i]/rounds[i],2))
        excelKPR.append(round(kills[i]/rounds[i],2))
        excelSRV.append(round(((rounds[i]-deaths[i])/rounds[i]),2))
        excelMKills.append(round(mk[i]/rounds[i],2))
        excelTrade.append(round(trade[i]/deaths[i],2))
        excelClutch.append(round(clutch[i]/rounds[i],2))
        excelPlants.append(round(plant[i]/(rounds[i]/2),2))
        excelHS.append(math.ceil(hs[i]))
        excelRound.append(rounds[i])
    #If match, update the value and replace at position
    else:
        excelHS[matchingValue] = math.ceil((excelHS[matchingValue]*excelKills[matchingValue] + hs[i]*kills[i])/(excelKills[matchingValue] + kills[i]))
        excelKills[matchingValue] = excelKills[matchingValue] + kills[i]
        excelDeaths[matchingValue] = excelDeaths[matchingValue] + deaths[i]
        excelKD[matchingValue] = round(excelKills[matchingValue]/excelDeaths[matchingValue],2)
        excelEntryKill[matchingValue] = excelEntryKill[matchingValue] + eKills[i]
        excelEntryDeath[matchingValue] = excelEntryDeath[matchingValue] + eDeaths[i]
        excelEntryPlusMinus[matchingValue] = excelEntryKill[matchingValue] - excelEntryDeath[matchingValue]
        excelKOST[matchingValue] = round((excelKOST[matchingValue]*excelRound[matchingValue] + kOST[i])/(excelRound[matchingValue] + rounds[i]),2)
        excelKPR[matchingValue] = round((excelKPR[matchingValue]*excelRound[matchingValue] + kills[i])/(excelRound[matchingValue] + rounds[i]),2)
        excelSRV[matchingValue] = round(((excelRound[matchingValue] + rounds[i]) - excelDeaths[matchingValue])/(excelRound[matchingValue] + rounds[i]),2)
        excelMKills[matchingValue] = round((excelMKills[matchingValue]*excelRound[matchingValue] + mk[i])/(excelRound[matchingValue]+rounds[i]),2)
        excelTrade[matchingValue] = round((excelTrade[matchingValue]*excelRound[matchingValue] + trade[i])/(excelRound[matchingValue]+rounds[i]),2)
        excelClutch[matchingValue] = round((excelClutch[matchingValue]*excelRound[matchingValue] + clutch[i])/(excelRound[matchingValue]+rounds[i]),2)
        excelPlants[matchingValue] = round((excelPlants[matchingValue]*excelRound[matchingValue] + plant[i])/(excelRound[matchingValue]+rounds[i]),2)
        
        excelRound[matchingValue] = excelRound[matchingValue] + rounds[i]
        excelRating[matchingValue] = round(ratingSys(excelKills[matchingValue], excelKD[matchingValue], excelMKills[matchingValue], excelEntryPlusMinus[matchingValue], excelPlants[matchingValue], excelClutch[matchingValue], excelKOST[matchingValue], excelSRV[matchingValue], excelRound[matchingValue]),2)
 



            



#Reput in the cell values now that they are updated
for i in range(len(excelUsername)):
    cell = excelMainSheet[excelCols[0] + str(i+2)]
    cell.value = excelUsername[i]

for i in range(len(excelRating)):
    cell = excelMainSheet[excelCols[1] + str(i+2)]
    cell.value = excelRating[i]

for i in range(len(excelKills)):
    cell = excelMainSheet[excelCols[2] + str(i+2)]
    cell.value = excelKills[i]

for i in range(len(excelDeaths)):
    cell = excelMainSheet[excelCols[3] + str(i+2)]
    cell.value = excelDeaths[i]

for i in range(len(excelKD)):
    cell = excelMainSheet[excelCols[4] + str(i+2)]
    cell.value = excelKD[i]

for i in range(len(excelEntryKill)):
    cell = excelMainSheet[excelCols[5] + str(i+2)]
    cell.value = excelEntryKill[i]

for i in range(len(excelEntryDeath)):
    cell = excelMainSheet[excelCols[6] + str(i+2)]
    cell.value = excelEntryDeath[i]

for i in range(len(excelEntryPlusMinus)):
    cell = excelMainSheet[excelCols[7] + str(i+2)]
    cell.value = excelEntryPlusMinus[i]

for i in range(len(excelKOST)):
    cell = excelMainSheet[excelCols[8] + str(i+2)]
    cell.value = excelKOST[i]

for i in range(len(excelKPR)):
    cell = excelMainSheet[excelCols[9] + str(i+2)]
    cell.value = excelKPR[i]

for i in range(len(excelSRV)):
    cell = excelMainSheet[excelCols[10] + str(i+2)]
    cell.value = excelSRV[i]

for i in range(len(excelMKills)):
    cell = excelMainSheet[excelCols[11] + str(i+2)]
    cell.value = excelMKills[i]

for i in range(len(excelTrade)):
    cell = excelMainSheet[excelCols[12] + str(i+2)]
    cell.value = excelTrade[i]

for i in range(len(excelClutch)):
    cell = excelMainSheet[excelCols[13] + str(i+2)]
    cell.value = excelClutch[i]

for i in range(len(excelPlants)):
    cell = excelMainSheet[excelCols[14] + str(i+2)]
    cell.value = excelPlants[i]

for i in range(len(excelHS)):
    cell = excelMainSheet[excelCols[15] + str(i+2)]
    cell.value = excelHS[i]

for i in range(len(excelRound)):
    cell = excelMainSheet[excelCols[16] + str(i+2)]
    cell.value = excelRound[i]

#what does this do?
for h in range(len(users)):
    temp = [users[h]]

# Save the changes
#This is how to get the Excel file to save the changes that were made
#workbook.save(filename=excelFile)


#These are examples of how to access the data in python per operator stats
#players[9].SingleOperatorStats('Flores')
#players[7].allOps()








