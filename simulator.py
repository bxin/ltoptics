class LaserTracker:
    def __init__(self, host, logger):
        self.position = [0, 0, 0]
        self.position_err = [0.001, 0.001, 0.001]

    def connect(self):
        pass

    def reconnect(self):
        pass

    def initialize(self):
        pass

    def measure(self):
        """
        returns position and associated standard deviation
        """
        return self.position, self.position_err

    def goto_position(self, position):
        self.position = position
