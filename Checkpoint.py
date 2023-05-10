import datetime
import time


class checkpoint:
    name : str
    number: str
    team : str
    mapReference : str
    distanceFromPrior : str
    timeAllowed : datetime
    firstTeamArrival : datetime
    firstTeamDeparture : datetime
    lastTeamDeparture : datetime
    
    def __init__(self,name,number):
        self.number = number
        self.name = name
       


    def show(self):
        # Print all the data
        print(self.number)
        print(self.name)
        print(self.team)
        print(self.mapReference)
        print(self.distanceFromPrior)
        print(self.timeAllowed)
        print(self.firstTeamArrival)
        print(self.firstTeamDeparture)

        