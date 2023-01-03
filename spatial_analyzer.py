import win32com.client
import pythoncom


class LaserTracker:
    def __init__(self, host, logger):
        print("__init__")
        import win32com.client
        self.client = win32com.client.Dispatch('SpatialAnalyzerSDK.Application')
        self.i_pt = -1
        self.log = logger
        self.instrument_id = None
        self.instrument_name = None
        self.last_point_name = f"{str(hash(str(self.i_pt)))[:6]}_AlignmentPt-{self.i_pt}"

        print("connect")
        self.log.info("Connecting to SpatialAnalyzer...")
        if not self.client.Connect("localhost"):
            raise ConnectionError

        self.log.info("getting last instrument")
        aux_i = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0.0)
        self.client.SetStep("Get Last Instrument Index")
        self.client.GetIntegerArg("Instrument ID", aux_i)

        #self.log.info("checking if instrument is located")
        #aux_i = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0.0)
        #self.client.SetStep("Get Last Instrument Index")
        #self.client.GetIntegerArg("Instrument ID", aux_i)

        # set units to millimeters
        self.client.SetStep("Set Active Units")
        self.client.SetDistanceUnitsArg("Length", "Millimeters")
        self.client.SetBoolArg("Display Inch Fractions?", False)
        self.client.SetDoubleArg("Inch Fraction Denominator?", 16.000000)
        self.client.SetBoolArg("Simplify Inch Fraction?", True)
        self.client.SetTemperatureUnitsArg("Temperature", "Celsius")
        self.client.SetAngularUnitsArg("Angular", "Degrees")
        self.client.ExecuteStep



    def reconnect(self):
        print("reconnect")
        pass

    def initialize(self):
        print("initialize")
        pass

    def measure(self):
        """
        returns position and associated standard deviation
        """
        print("measure")
        # measure point
        self.client.SetStep("Measure Single Point Here")
        self.client.SetColInstIdArg("Instrument ID", "", 0)
        self.client.SetPointNameArg("Target ID", "A", "LTOptics", self.last_point_name)
        self.client.SetBoolArg("Measure Immediately", True)
        self.client.SetFilePathArg("HTML Prompt File (optional)", "", False)
        self.client.ExecuteStep

        # get point value
        self.client.SetStep("Get Point Coordinate")
        self.client.SetPointNameArg("Point Name", "A", "LTOptics", self.last_point_name)
        self.client.ExecuteStep

        x = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
        y = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
        z = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
        self.client.GetVectorArg("Vector Representation", x, y, z)

        # stdev:
        # NrkSdk.SetStep("Create Point Uncertainty Fields")
        # 	Dim ptNameList(1)
        # 	ptNameList( 0 ) = "A::LTOptics::426745_AlignmentPt-1"
        # 	Dim vPointObjectList As Object = New System.Runtime.InteropServices.VariantWrapper(ptNameList)
        # 	NrkSdk.SetPointNameRefListArg("Point Name List", vPointObjectList)
        # 	NrkSdk.SetIntegerArg("Number of Samples", 1000)
        # NrkSdk.ExecuteStep( )
        #
        self.client.SetStep("Get Point Properties")
        self.client.SetPointNameArg("Point Name", "A", "LTOptics", self.last_point_name)
        self.client.ExecuteStep
        #
        # Dim value As Double
        # NrkSdk.GetDoubleArg("Planar Offset", value)
        #
        # Dim value As Double
        # NrkSdk.GetDoubleArg("Radial Offset", value)
        #
        # Dim value As Double
        # NrkSdk.GetDoubleArg("Ux", value)
        #
        # Dim value As Double
        # NrkSdk.GetDoubleArg("Uy", value)
        #
        # Dim value As Double
        # NrkSdk.GetDoubleArg("Uz", value)
        #
        # Dim value As Double
        # NrkSdk.GetDoubleArg("Umag", value)
        #
        # Dim bUseHighX As Boolean
        # Dim highTolX As Double
        # Dim bUseHighY As Boolean
        # Dim highTolY As Double
        # Dim bUseHighZ As Boolean
        # Dim highTolZ As Double
        # Dim bUseHighM As Boolean
        # Dim highTolM As Double
        # Dim bUseLowX As Boolean
        # Dim lowTolX As Double
        # Dim bUseLowY As Boolean
        # Dim lowTolY As Double
        # Dim bUseLowZ As Boolean
        # Dim lowTolZ As Double
        # Dim bUseLowM As Boolean
        # Dim lowTolM As Double
        # NrkSdk.GetToleranceVectorOptionsArg("Position Tolerance",
        # 	bUseHighX, highTolX, bUseHighY, highTolY,
        # 	bUseHighZ, highTolZ, bUseHighM, highTolM,
        # 	bUseLowX, lowTolX, bUseLowY, lowTolY,
        # 	bUseLowZ, lowTolZ, bUseLowM, lowTolM)
        #
        # Dim xVal As Double
        # Dim yVal As Double
        # Dim zVal As Double
        # NrkSdk.GetVectorArg("Component Weights", xVal, yVal, zVal)
        #


        return [x.value, y.value, z.value], [0, 0, 0]  # todo stdev!

    def goto_position(self, position):
        """
        input: position in x,y,z coordinates
        """
        print("goto")
        # update current point name
        self.i_pt += 1
        self.last_point_name = f"{str(hash(str(self.i_pt)))[:6]}_AlignmentPt-{self.i_pt}"

        # create a point
        self.client.SetStep("Construct a Point in Working Coordinates")
        self.client.SetVectorArg("Working Coordinates", float(position[0]), float(position[1]), float(position[2]))
        self.client.SetPointNameArg("Point Name", "A", "LTOptics", self.last_point_name)
        self.client.ExecuteStep

        # point
        self.client.SetStep("Point At Target")
        self.client.SetColInstIdArg("Instrument ID", "", 0)
        self.client.SetPointNameArg("Target ID", "A", "LTOptics", self.last_point_name)
        self.client.SetFilePathArg("HTML Prompt File (optional)", "", False)
        self.client.ExecuteStep

        self.log.info(f"Goto position: {position}, point {self.i_pt}")
