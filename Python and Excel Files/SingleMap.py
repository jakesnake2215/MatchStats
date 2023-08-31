import numpy as np
import statistics
import Operators
#Parses json file for all stats
#Unable to determine how the disable works !!NEEDS TESTING!!

def singleMap(dict):
    #define all variables to pass through/use
    Team1Score = 0
    Team2Score = 0
    #Takes the team names from the first round
    Team1Name = dict["rounds"][0]["teams"][0]["name"]
    Team2Name = dict["rounds"][0]["teams"][1]["name"]
    UsernameList = []
    UsernameLookup = {}
    KAmount = []
    DAmount = []
    HSPercent = []
    RoundCount = []
    AtkMain = []
    DefMain = []
    #Note: The operator arrays are filled to -1, this is because if it was initialized at 0, could show a false positive of mute being played because he is operator 0
    RoundOps = np.full(10,-1)
    DTotalOps = np.array([])
    ATotalOps = np.array([])
    ARoundOps = np.full(10,-1)
    DRoundOps = np.full(10,-1)
    AMainOp = np.array([])
    DMainOp = np.array([])
    WinningTeamMembers = []
    EntryKills = np.zeros(10)
    EntryDeaths = np.zeros(10)
    Plants = np.zeros(10)
    Defusal = np.zeros(10)
    KOSTRounds = np.zeros(10)
    KOSTTotal=np.zeros(10)
    KOSTSurv = np.zeros(10)
    Clutches = np.zeros(10)
    Multikills = np.zeros(10)
    Trades = np.zeros(10)
    OpKills = np.zeros((10,68))
    OpDeaths = np.zeros((10,68))
    OpEKills = np.zeros((10,68))
    OpEDeaths = np.zeros((10,68))
    OpHS = np.zeros((10,68))
    OpPlants = np.zeros((10,68))
    OpDefusal = np.zeros((10,68))
    OpKOST = np.zeros((10,68))
    OpKOSTRound = np.zeros((10,68))
    OpClutches = np.zeros((10,68))
    OpMultikills = np.zeros((10,68))
    OpTrades = np.zeros((10,68))
    OpRounds = np.zeros((10,68))
    #Takes the first round map name to output
    Map = dict["rounds"][0]["map"]["name"]
    #takes the total number of kills in the map and hs percent and rounds played
    for i in range(10):
        #Appends list at the end of rounds to track the basic stats that is given, could be improved if needed
        UsernameList.append(dict["stats"][i]["username"])
        KAmount.append(dict["stats"][i]["kills"])
        DAmount.append(dict["stats"][i]["deaths"])
        HSPercent.append(dict["stats"][i]["headshotPercentage"])
        RoundCount.append(dict["stats"][i]["rounds"])
        #creates a lookup table for each player, 0-9 based on name and how incrememented in the json file
        #Makes it easier to upload names to further stats, because can organize players
        UsernameLookup[UsernameList[i]] = i
    #large loop to look at all rounds
    #!! Restructing Note, should do a double for loop where outer loop is rounds, and inner loops in actions in rounds !!
    for i in range(len(dict["rounds"])):
        j = 0
        #actions is number of occurences in the 'main phase' of each round, contains everything that a player can do to impact each round
        #!!2 Different logics between while and for loop, should restructure!!
        
        #Updates team score each round, so the final score is output at the end of all rounds
        Team1Score = dict["rounds"][i]["teams"][0]["score"]
        Team2Score = dict["rounds"][i]["teams"][1]["score"]
        actions = len(dict["rounds"][i]["matchFeedback"])
        #loops through all the actions looking for the first kill to occur
        for v in range(len(dict["rounds"][i]["players"])):
            
            #if the operator that is selected by the player is > 33 it is an attacker by my dictionary, otherwise its a defender
            #if an attacker, needs to check if a repick occurs
            if(Operators.Operators[dict["rounds"][i]["players"][v]["operator"]["name"]] > 33):
                #Puts the Player on the operator that they played that round in their 'spot' in the array, and given a numerical value from the dict 'Operators'
                RoundOps[UsernameLookup[dict["rounds"][i]["players"][v]["username"]]] = Operators.Operators[dict["rounds"][i]["players"][v]["operator"]["name"]]
                #as attackers can swap in prep phase, look for the match feedback for an operator swap and update the value
                #Checks through all the round actions to see if an operator is swapped off
                for c in range(actions):
                    if(dict["rounds"][i]["matchFeedback"][c]["type"]["name"] == "OperatorSwap"):
                        #Updates the value in similar value
                        RoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][c]["username"]]] = Operators.Operators[dict["rounds"][i]["matchFeedback"][c]["operator"]["name"]]
                        
            else:
                #Defenders do not change, so the first time seen can be set
                RoundOps[UsernameLookup[dict["rounds"][i]["players"][v]["username"]]] = Operators.Operators[dict["rounds"][i]["players"][v]["operator"]["name"]]
            #Hard to read
            #Puts all the operator plays into the 2d array for the Operator Stats Page, Puts the username lookup and and operator location and increments the play count
            OpRounds[UsernameLookup[dict["rounds"][i]["players"][v]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["players"][v]["username"]]]] = OpRounds[UsernameLookup[dict["rounds"][i]["players"][v]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["players"][v]["username"]]]] + 1
        
        #Loops for opening kill
        while(dict["rounds"][i]["matchFeedback"][j]["type"]["name"] != "Kill"):
            #Will cause a runtime error if last loop and final thing is a kill
            j = j+1
            #edgecase that no kills occur
            #Further Testing needed on if this works
            if(j == actions):
                break
        #when the first kill occurs give a kill and death to user and target respectively
        #also added to individual op per player deaths 
        if(j<actions):
            EntryKills[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] = EntryKills[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] + 1
            EntryDeaths[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]] = EntryDeaths[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]] + 1
            OpEKills[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]]] = OpEKills[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]]] + 1
            OpEDeaths[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]],RoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]]] = OpEDeaths[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]],RoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["target"]]]] + 1
        j=0
        
        #looks for the plant to go down
        while(dict["rounds"][i]["matchFeedback"][j]["type"]["name"] != "DefuserPlantComplete"):
            j = j+1
            if(j == actions):
                break
        #gives the planter a point if they get defuser down and what operator they played
        if(j<actions):
            Plants[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] = Plants[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]] + 1
            OpPlants[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]]] = OpPlants[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][j]["username"]]]]+1

        #tracks KOST (Kills, Objectives[plants], Survival, Trades)
        KOSTRounds = np.zeros(10)
        #If did nothing but lived, should have 1
        KOSTSurv = np.ones(10)
        OpKOSTRound = np.zeros((10,Operators.OpNumbers))
        #looks for survival rate and trades
        for k in range(len(dict["rounds"][i]["matchFeedback"])):
            if(dict["rounds"][i]["matchFeedback"][k]["type"]["name"] == "Kill"):
                #If a kill happens, the round is counted for KOST, but not survival
                KOSTRounds[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["username"]]] = 1
                KOSTSurv[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["target"]]] = 0
                #looks for trades, when a user dies, capture tOD and who killed them
                #then iterate through until the time is past the time to trade to see if the trade occurs
                timeOfDeath = dict["rounds"][i]["matchFeedback"][k]["timeInSeconds"]
                UserToBeTraded = dict["rounds"][i]["matchFeedback"][k]["username"]
                l=0
                #If the time happens, loops through all actions to see if it is within the 10 second trade time, otherwise go to next action
                while(timeOfDeath-Operators.timeToTrade < dict["rounds"][i]["matchFeedback"][k+l]["timeInSeconds"]):
                    if(dict["rounds"][i]["matchFeedback"][k+l]["type"]["name"] == "Kill" and dict["rounds"][i]["matchFeedback"][k+l]["target"] == UserToBeTraded):
                        KOSTRounds[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["target"]]] = 1
                        Trades[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["target"]]] = Trades[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["target"]]] + 1
                        OpTrades[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["target"]],RoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["target"]]]] = OpTrades[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["target"]],RoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["target"]]]] + 1
                        break
                    else:
                        l = l+1
                        if(k+l == actions):
                            break
                
            #adds plant if get plant down
            elif(dict["rounds"][i]["matchFeedback"][k]["type"]["name"] == "DefuserPlantComplete"):
                KOSTRounds[UsernameLookup[dict["rounds"][i]["matchFeedback"][k]["username"]]] = 1
        #adds a round where you add to your kost, makes sure not to duplicate if 2 things occur in a round
        #Adds rounds to if they survived
        KOSTRounds = KOSTRounds + KOSTSurv
        
        #If higher than one, sets back to 1 to not overrate the KOST rounds if someone does more than one KOST action in a round
        for n in range(len(KOSTRounds)):
            if(KOSTRounds[n] > 1):
                KOSTRounds[n] = 1
            OpKOSTRound[n, RoundOps[n]] = KOSTRounds[n]
        #Adds the KOST to total number of rounds that players achieved during the number of rounds
        KOSTTotal = KOSTTotal + KOSTRounds
        OpKOST = OpKOST + OpKOSTRound

        #tracks each user for their multikills over a game
        #Checks for HS percentage and Kills for operators per round, also looks to see if number of kills was greater than 1
        for g in range(len(dict["rounds"][i]["stats"])):
            
            OpKills[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] = OpKills[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] + dict["rounds"][i]["stats"][g]["kills"]
            OpHS[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] = OpHS[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] + dict["rounds"][i]["stats"][g]["headshots"]
            if(dict["rounds"][i]["stats"][g]["died"] == True):
                OpDeaths[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] = OpDeaths[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]],RoundOps[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] + 1
            if(dict["rounds"][i]["stats"][g]["kills"] > 1):
                Multikills[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]] = Multikills[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]] + 1
                OpMultikills[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]], RoundOps[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] = OpMultikills[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]], RoundOps[UsernameLookup[dict["rounds"][i]["stats"][g]["username"]]]] + 1
        
        #tracks clutching based on if your team won and you were the only player on your team alive
        if(dict["rounds"][i]["teams"][0]["won"] == True):
            WinningTeam = 0
        else:
            WinningTeam = 1
        
        #start with 5 players alive on the winning team, if a winning team member dies, then reduce the number alive
        #add winning round members
        WinningTeamMembers = []
        ClutchPlayer = ''
        for m in range(len(dict["rounds"][i]["players"])):
            if(dict["rounds"][i]["players"][m]["teamIndex"] == WinningTeam):
                WinningTeamMembers.append(dict["rounds"][i]["players"][m]["username"])
        ClutchAlive = 5
        #if the winning team member died, then reduce the number, else assign that person the clutch player, if nobody is alive on a team then this doesnt get added to clutch total
        for a in range(len(WinningTeamMembers)):
            for b in range(len(dict["rounds"][i]["stats"])):
                if(WinningTeamMembers[a] == dict["rounds"][i]["stats"][b]["username"] and dict["rounds"][i]["stats"][b]["died"] == True):
                    ClutchAlive = ClutchAlive - 1
                else:
                    ClutchPlayer = dict["rounds"][i]["stats"][b]["username"]
        #Check again to see that only 1 player was alive on the winning team to confirm that it is a clutch and give it to player and the operator they played that round
        if(ClutchAlive == 1):
            Clutches[UsernameLookup[ClutchPlayer]] = Clutches[UsernameLookup[ClutchPlayer]] + 1
            OpClutches[UsernameLookup[ClutchPlayer],RoundOps[UsernameLookup[ClutchPlayer]]] = OpClutches[UsernameLookup[ClutchPlayer],RoundOps[UsernameLookup[ClutchPlayer]]] + 1

        #find most played operator
        ARoundOps = np.full(10,-1)
        DRoundOps = np.full(10,-1)
        #look through all players
        for v in range(len(dict["rounds"][i]["players"])):
            
            #if the operator that is selected by the player is > 33 it is an attacker by my dictionary, otherwise its a defender
            if(Operators.Operators[dict["rounds"][i]["players"][v]["operator"]["name"]] > 33):
                ARoundOps[UsernameLookup[dict["rounds"][i]["players"][v]["username"]]] = Operators.Operators[dict["rounds"][i]["players"][v]["operator"]["name"]]
                
                #as attackers can swap in prep phase, look for the match feedback for an operator swap and update the value
                for c in range(actions):
                    if(dict["rounds"][i]["matchFeedback"][c]["type"]["name"] == "OperatorSwap"):
                        ARoundOps[UsernameLookup[dict["rounds"][i]["matchFeedback"][c]["username"]]] = Operators.Operators[dict["rounds"][i]["matchFeedback"][c]["operator"]["name"]]
            #defenders will be the same            
            else:
                DRoundOps[UsernameLookup[dict["rounds"][i]["players"][v]["username"]]] = Operators.Operators[dict["rounds"][i]["players"][v]["operator"]["name"]]
                
        #use the dict defined numerical value of the op to add to a total array
        ATotalOps = np.append(ATotalOps, ARoundOps)
        DTotalOps = np.append(DTotalOps, DRoundOps)
        #find all main ops by using array math to loop through and find each players most played operator
    for z in range(len(dict["rounds"][1]["players"])):
        AMainOp = np.array([])
        DMainOp = np.array([])
        #Separate players into their individual operators to check for main operator that a single person
        for q in range(len(dict["rounds"])):
            AMainOp = np.append(AMainOp, ATotalOps[q*10 + z])
            AMainOp = AMainOp[AMainOp>=0]
            DMainOp = np.append(DMainOp, DTotalOps[q*10 + z])
            DMainOp = DMainOp[DMainOp>=0]
        #add each users main op back to main array
        #Create an array to check what the main operator each person played
        AMain = statistics.mode(AMainOp)
        AMain = int(AMain)
        DMain = statistics.mode(DMainOp)
        AtkMain.append(Operators.OperatorsValues[AMain])
        DefMain.append(Operators.OperatorsValues[DMain])
    
        
    #all return values are array values of # of players
    return [UsernameList,KAmount,DAmount,EntryKills,EntryDeaths,KOSTTotal, HSPercent,Multikills,Trades, Clutches, Plants,Defusal,AtkMain,DefMain,RoundCount, OpKills, OpDeaths, OpEKills, OpEDeaths, OpKOST, OpHS, OpMultikills, OpTrades, OpClutches, OpPlants,OpDefusal, OpRounds, Map, Team1Score, Team2Score, Team1Name, Team2Name]
   