import numpy as np
class BasicStats:
    #Initializes variables, (Why do I have to do this?)
    def __init__(self, Username, Kills, Deaths, KD, EKills, EDeaths, Entry, KOST, KPR, SRV, MKills, Trade, Clutch, Plants, Defuse, HSPercent, FavAtk, FavDef, Rounds, OpK, OpD, OpKD, OpEK, OpED, OpEntry, OpKOST, OpKPR, OpSRV, OpMKills, OpTrade, OpClutch, OpPlants, OpDefuse, OpHS, OpRounds, MapPlayed, Team1Score, Team2Score, Team1Name, Team2Name):
        self.Username = Username
        self.Kills = Kills
        self.Deaths = Deaths
        self.KD = KD
        self.EKills = EKills
        self.EDeaths = EDeaths
        self.Entry = Entry
        self.KOST = KOST
        self.KPR = KPR
        self.SRV = SRV
        self.MKills = MKills
        self.Trade = Trade
        self.Clutch = Clutch
        self.Plants = Plants
        self.Defuse = Defuse
        self.HSPercent = HSPercent
        self.FavAtk = FavAtk
        self.FavDef = FavDef
        self.Rounds = Rounds
        self.OpK = OpK
        self.OpD = OpD
        self.OpKD = OpKD
        self.OpEK = OpEK
        self.OpED = OpED
        self.OpEntry = OpEntry
        self.OpKOST = OpKOST
        self.OpKPR = OpKPR
        self.OpSRV = OpSRV
        self.OpMKills = OpMKills
        self.OpTrade = OpTrade
        self.OpClutch = OpClutch
        self.OpPlants = OpPlants
        self.OpDefuse = OpDefuse
        self.OpHS = OpHS
        self.OpRounds = OpRounds
        self.MapPlayed = MapPlayed
        self.Team1Score = Team1Score
        self.Team2Score = Team2Score
        self.Team1Name = Team1Name
        self.Team2Name = Team2Name
        
#Rating System, very basic
#Same rating system as before
    def rating(self):
        rating = (self.KD**2 + 0.4*(self.Kills) + 0.15*self.MKills)/self.Rounds + 0.75*(self.Entry)/self.Rounds + (self.Plants + self.Clutch)/self.Rounds + self.KOST + self.SRV/3
        return rating
#Prints out all 'relevant' stats, in similar format as siegeGG, adds multikills and trades for greater visibility
    def printIndivStat(self, intro):
        #defines the rating in this def
        rating = self.rating()
        #if the first user printed, prints a header
        
        #Top part of the print, gives the map, and score and header
        if(intro == 1):
            print('Map: ' + self.MapPlayed)
            print('')
            print(self.Team1Name + ' - ' + self.Team2Name)
            print(str(self.Team1Score) + '-' + str(self.Team2Score))
            formatted_string = "{:<15} | {:<6} | {:<10} | {:<8} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<10} | {:<10}".format(
                'Username', 'Rating', 'K-D(KD)', 'Entry', 'KOST', 'KPR', 'SRV', '1vX', 'Plants', 'Multis', 'Trades', 'HS', 'Attacker', 'Defender', '', ''
            )
            #copies the underlined text to the length of the text above and creates a line
            underlined_string = formatted_string + '\n' + '-' * len(formatted_string)
            print(underlined_string)

        #formats the plus or minus in front of the KD and entry to make it + or -
        PlusMinus = self.Kills - self.Deaths
        if(PlusMinus > 0):
            strKD = str(self.Kills) + '-' + str(self.Deaths) + '(+' + str(PlusMinus)+')'
        else:
            strKD = str(self.Kills) + '-' + str(self.Deaths) + '(' + str(PlusMinus)+')'
        
        #Same formatting for Entry Stats
        EPlusMinus = self.EKills - self.EDeaths
        if(EPlusMinus > 0):
            strEntry = str(int(self.EKills)) + '-' + str(int(self.EDeaths)) + '(+'+str(int(EPlusMinus))+')'
        else:
            strEntry = str(int(self.EKills)) + '-' + str(int(self.EDeaths)) + '('+str(int(EPlusMinus))+')'
        #formatting the text for each user and prints
        formatKOST = "{:.2f}".format(self.KOST)
        formatKPR = "{:.2f}".format(self.KPR)
        formatRating = "{:.2f}".format(rating)
        formatted_string = "{:<15} | {:<6} | {:<10} | {:<8} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<10} | {:<10}".format(
            self.Username,
            formatRating,
            strKD,
            strEntry,
            formatKOST,
            formatKPR,
            str(int(self.SRV*100)) + '%',
            int(self.Clutch),
            int(self.Plants),
            int(self.MKills),
            int(self.Trade),
            str(int(self.HSPercent)) + '%',
            self.FavAtk,
            self.FavDef
        )
        print(formatted_string)
    #Define ratings for Operator Ratings for a player
    def OperatorRating(self):
        OperatorRating = np.zeros(OpNumbers)
        #rating = (1.5*self.KD + 0.25*(self.Kills) + 0.15*self.MKills)/self.Rounds + 0.75*(self.Entry)/self.Rounds + (self.Plants + self.Clutch)/self.Rounds + self.KOST + self.SRV/3
        for j in (range(OpNumbers)):
            if self.OpRounds[j] == 0:
                OperatorRating[j] = 0
            else:
                OperatorRating[j] = ((self.OpKD[j]**2 + 0.4*(self.OpK[j]) + 0.15*self.OpMKills[j])/self.OpRounds[j] + 0.75*(self.OpEntry[j])/self.OpRounds[j] + (self.OpPlants[j] + self.OpClutch[j])/self.OpRounds[j] + self.OpKOST[j] + (self.OpSRV[j])/3)
        return OperatorRating
    #Simple Way to read all player Operator Rating, just in python currently, but could be phased out
    def AllOps(self):
        Op = self.OperatorRating()
        for k in (range(OpNumbers)):
            number_str = Op[k]
            roundedRating = "{:.2f}".format(float(number_str))
            print(OperatorsValues[k] + ': ' + roundedRating)
    #Similar to above, can look at individual player and full rating for a single operator, python only, either phase out or can be used in maybe a different aspect
    #Very similar formatting to the full list for a single map, should be merged
    def SingleOperatorStats(self, inputStr):
        OperatorValue = Operators[inputStr]
        print('\n')
        OpsRate = self.OperatorRating()
        
        formatted_string = "{:<12} | {:<10} | {:<6} | {:<10} | {:<8} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6}".format(
                'Player','Op Name', 'Rating', 'K-D(KD)', 'Entry', 'KOST', 'KPR', 'SRV', '1vX', 'Plants', 'Multis', 'Trades', 'HS', 'Rounds', '')
        underlined_string = formatted_string + '\n' + '-' * len(formatted_string)
        print(underlined_string)
        PlusMinus = self.OpK[OperatorValue] - self.OpD[OperatorValue]
        if(PlusMinus > 0):
            strKD = str(int(self.OpK[OperatorValue])) + '-' + str(int(self.OpD[OperatorValue])) + '(+' + str(int(PlusMinus))+')'
        else:
            strKD = str(int(self.OpK[OperatorValue])) + '-' + str(int(self.OpD[OperatorValue])) + '(' + str(int(PlusMinus))+')'
        
        EPlusMinus = self.OpEK[OperatorValue] - self.OpED[OperatorValue]
        if(EPlusMinus > 0):
            strEntry = str(int(self.OpEK[OperatorValue])) + '-' + str(int(self.OpED[OperatorValue])) + '(+'+str(int(EPlusMinus))+')'
        else:
            strEntry = str(int(self.OpEK[OperatorValue])) + '-' + str(int(self.OpED[OperatorValue])) + '('+str(int(EPlusMinus))+')'
        #formatting the text for each user and prints
        if self.OpK[OperatorValue] == 0:
            HS = 0
        else:
            HS = self.OpHS[OperatorValue]/self.OpK[OperatorValue]*100
        formatKOST = "{:.2f}".format(self.OpKOST[OperatorValue])
        formatKPR = "{:.2f}".format(self.OpKPR[OperatorValue])
        formatRating = "{:.2f}".format(OpsRate[OperatorValue])
        
        formatted_string = "{:<12} | {:<10} | {:<6} | {:<10} | {:<8} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6} | {:<6}".format(
            self.Username,
            inputStr,
            formatRating,
            strKD,
            strEntry,
            formatKOST,
            formatKPR,
            str(int(self.OpSRV[OperatorValue]*100)) + '%',
            int(self.OpClutch[OperatorValue]),
            int(self.OpPlants[OperatorValue]),
            int(self.OpMKills[OperatorValue]),
            int(self.OpTrade[OperatorValue]),
            str(int(HS)) + '%',
            str(int(self.OpRounds[OperatorValue]))
        )
        print(formatted_string)