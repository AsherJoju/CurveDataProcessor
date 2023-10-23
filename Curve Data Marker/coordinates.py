class Coordinates:
    
    def __init__(self, latitude: float, longitude: float, index = "") -> None:
        self.latitude = latitude
        self.longitude = longitude
        self.index = index
    
    
    def __repr__(self) -> str:
        return f"{'Curve No. ' + self.index + ', ' if self.index else ''}{self.latitude}, {self.longitude}"
    
    
    def distance_to(self, coordinates) -> float:
        if not isinstance(coordinates, Coordinates):
            raise TypeError("The parameter 'coordinates' must be of type 'Coordinates'")
        
        distance = (
            (self.latitude - coordinates.latitude) ** 2 +
            (self.longitude - coordinates.longitude) ** 2
        ) ** (1/2) # Square Root
        
        return distance
