import win32com.client
client = win32com.client.Dispatch('SpatialAnalyzerSDK.Application')

client.SetStep("Construct a Point in Working Coordinates")
client.SetPointNameArg("Point Name", "a", "b", "c")
client.SetVectorArg("Working Coordinates", 0.000000, 0.000000, 0.000000)
client.ExecuteStep()

client.SetStep("Point At Target")
client.SetColInstIdArg("Instrument ID", "0-API Radian (Live)")
client.SetPointNameArg("Target ID", "", "", "")
client.SetFilePathArg("HTML Prompt File (optional)", "", False)
client.ExecuteStep

client.SetStep("Measure Single Point Here")
client.SetColInstIdArg("Instrument ID", "", "0-API Radian (Live)")
client.SetPointNameArg("Target ID", "", "", "")
client.SetBoolArg("Measure Immediately", True)
client.SetFilePathArg("HTML Prompt File (optional)", "", False)
client.ExecuteStep()

client.SetStep("Measure Single Point Here")
# client.SetColInstIdArg("Instrument ID", "", 0-API Radian (Live))
client.SetPointNameArg("Target ID", "A", "Main", "test")
client.SetBoolArg("Measure Immediately", True)
client.SetFilePathArg("HTML Prompt File (optional)", "", False)
client.ExecuteStep()

# xxxx


#  connect to the server
if not client.Connect("localhost"):
    raise ConnectionError

# create a point
client.SetStep("Construct a Point in Working Coordinates")
client.SetVectorArg("Working Coordinates", 10.1234, 20.2345, 30.5678)
client.SetPointNameArg("Point Name", "", "TestGrp", "TestPt3")
client.ExecuteStep

# todo:
# client.GetMPStepResult

# point
client.SetStep("Point At Target")
client.SetColInstIdArg("Instrument ID", "", 0)
client.SetPointNameArg("Target ID", "A", "TestGrp", "TestPt3")
client.SetFilePathArg("HTML Prompt File (optional)", "", False)
client.ExecuteStep

# measure point
client.SetStep("Measure Single Point Here")
client.SetColInstIdArg("Instrument ID", "", 0)
client.SetPointNameArg("Target ID", "A", "TestGrp", "TestPt3")
client.SetBoolArg("Measure Immediately", True)
client.SetFilePathArg("HTML Prompt File (optional)", "", False)
client.ExecuteStep

# xxxx

# create a point
client.SetStep("Construct a Point in Working Coordinates")
client.SetVectorArg("Working Coordinates", 0, 0, 0)
client.SetPointNameArg("Point Name", "", "TestGrp", "TestPt4")
client.ExecuteStep

# todo:
# client.GetMPStepResult

# point
client.SetStep("Point At Target")
client.SetColInstIdArg("Instrument ID", "", 0)
client.SetPointNameArg("Target ID", "A", "TestGrp", "TestPt4")
client.SetFilePathArg("HTML Prompt File (optional)", "", False)
client.ExecuteStep

# measure point
client.SetStep("Measure Single Point Here")
client.SetColInstIdArg("Instrument ID", "", 0)
client.SetPointNameArg("Target ID", "A", "TestGrp", "TestPt4")
client.SetBoolArg("Measure Immediately", True)
client.SetFilePathArg("HTML Prompt File (optional)", "", False)
client.ExecuteStep

# delete all measurements
# Delete Measurements:


# point and measure:
# Measure Existing Single Point


# get point data
# Coordinate
client.SetStep("Get Point Coordinate")
client.SetPointNameArg("Point Name", "A", "TestGrp", "TestPt3")
client.ExecuteStep

# DOUBLE
# xVal;
# DOUBLE
# yVal;
# DOUBLE
# zVal;
x = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
y = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
z = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
client.GetVectorArg("Vector Representation", x, y, z)

# DOUBLE
# value;
# client.GetDoubleArg("X Value", & value);
#
# DOUBLE
# value;
# client.GetDoubleArg("Y Value", & value);https://public.acontis.com/manuals/EC-Monitor/3.1/html/ras.html
#
# DOUBLE
# value;
# client.GetDoubleArg("Z Value", & value);


# get point

client.SetStep("Get Current Instrument Position Update")
client.SetColInstIdArg("Instrument ID", "", 0)
client.SetStringArg("Reporting Frame", "Instrument Base")
client.SetBoolArg("Polar Coordinates?", False)
x = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
y = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
z = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
client.ExecuteStep
client.GetDoubleArg("X / R", x)
client.GetDoubleArg("Y / Theta (Degrees)", y)
client.GetDoubleArg("Z / Phi (Degrees)", z)
print(x.value, y.value, z.value)

# Dim bSendStatus As Boolean
# bSendStatus = client.ExecuteStep()


from win32com.client import gencache

mod = gencache.GetModuleForProgID("SpatialAnalyzerSDK.Application")
app = mod.Application()

# client.SetStep("Set OPC DA Tag Value Integer")
# client.SetStringArg("OPC Server DA Tag Name", "test_set")
# client.SetIntegerArg("Value", 299)
# client.ExecuteStep()
#
# client.SetStep("Make a Point Name Ref List From a Group")
# client.SetCollectionObjectNameArg("Group Name", "NOMINAL", "SMR CENTERS")
# client.ExecuteStep()
#
# CStringArray
# ptNameList
# SDKHelper
# helper(client)
# helper.GetPointNameRefListArgHelper("Resultant Point Name List", ptNameList)
#
# client.SetStep("Get i-th Point Name From Point Name Ref List (Iterator)")
# CStringArray
# ptNameList
# SDKHelper
# helper(client)
# helper.SetPointNameRefListArgHelper("Point Name List", ptNameList)
# client.SetIntegerArg("Point Name Index", 0)
# client.NOT_SUPPORTED("Step to Jump at End of List")
# client.ExecuteStep()
#
# BSTR
# sValue = NULL
# client.GetStringArg("Collection", & sValue)
# CString
# name = sValue
# ::SysFreeString(sValue)
#
# BSTR
# sValue = NULL
# client.GetStringArg("Group", & sValue)
# CString
# name = sValue
# ::SysFreeString(sValue)
#
# BSTR
# sValue = NULL
# client.GetStringArg("Target", & sValue)
# CString
# name = sValue
# ::SysFreeString(sValue)
#
# BSTR
# sCol = NULL
# BSTR
# sGrp = NULL
# BSTR
# sTarg = NULL
# client.GetPointNameArg("Resulting Point Name", & sCol, & sGrp, & sTarg)
# CString
# collection = sCol
# CString
# group = sGrp
# CString
# target = sTarg
# ::SysFreeString(sCol)
# ::SysFreeString(sGrp)
# ::SysFreeString(sTarg)
#
# client.SetStep("Point At Target")
# client.SetColInstIdArg("Instrument ID", "", 0 - API
# Radian(Live))
# client.SetPointNameArg("Target ID", "A", "Main", "test")
# client.SetFilePathArg("HTML Prompt File (optional)", "", FALSE)
# client.ExecuteStep()
#
# client.SetStep("Measure Single Point Here")
# client.SetColInstIdArg("Instrument ID", "", 0 - API
# Radian(Live))
# client.SetPointNameArg("Target ID", "A", "Main", "test")
# client.SetBoolArg("Measure Immediately", TRUE)
# client.SetFilePathArg("HTML Prompt File (optional)", "", FALSE)
# client.ExecuteStep()
#
# client.SetStep("Jump To Step")
# client.NOT_SUPPORTED("Step to Jump To")
# client.ExecuteStep()
#
# client.SetStep("Exit Measurement Plan")
# client.ExecuteStep()
#
