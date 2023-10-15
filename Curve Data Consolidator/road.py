from enum import Enum


class RoadState(Enum):

    NONE = 0
    TANGENT = 1
    CV = 2
    CURVE = 3


class ValueList:
    
    def __init__(self):
        self.values = []
    
    
    def max(self):
        values = []
        
        for value in self.values:
            if value != None:
                values.append(value)
        
        try:
            return max(values)
        except:
            return None
    
    
    def min(self):
        values = []
        
        for value in self.values:
            if value != None:
                values.append(value)
        
        try:
            return min(values)
        except:
            return None
    
    
    def mean(self):
        values = []
        
        for value in self.values:
            if value != None:
                values.append(value)
        
        try:
            return sum(values) / len(values)
        except:
            return None
    
    
    def mode(self):
        values = []
        
        for value in self.values:
            if value != None:
                values.append(value)
        
        try:
            return max(values, key=values.count)
        except:
            return None


class Road:
    
    def __init__(self):
        self.speed_list = ValueList()
        self.hr_list = ValueList()
        self.gsr_list = ValueList()
    
    
    def get_stable_values(self):
        stable_values = Road()
        
        avr_spd = self.speed_list.mean()
        
        if avr_spd == None:
            return stable_values
        
        # try:
        if True:
            ind = 0
        
            for i in range(len(self.speed_list.values)):
                if self.speed_list.values[i] >= avr_spd:
                    ind = i
                    break
            
            for i in range(ind, len(self.speed_list.values)):
                stable_values.speed_list.values.append(self.speed_list.values[i])
                stable_values.hr_list.values.append(self.hr_list.values[i])
                stable_values.gsr_list.values.append(self.gsr_list.values[i])
            
            return stable_values
        # except:
        #     return stable_values
