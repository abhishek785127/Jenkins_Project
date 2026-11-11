
import win32com.client as win32
import time
import os
from datetime import datetime
from win32com.client import gencache

# Initialize CANoe Application
CANoe = win32.DispatchEx("CANoe.Application")
time.sleep(10)

# Open CANoe configuration
CANoe.Open(r"F:\Jenkins\Jenkins\Jenkins\ACC_RBS\pyrbs.cfg")
time.sleep(5)

# Test configuration details
# Generate a unique name for every execution
testConfigName = "AutomatedTests_VT_" + datetime.now().strftime("%Y%m%d_%H%M%S")

testUnitsFolder = r"F:\Jenkins\Jenkins\Jenkins\vTEST_Project\ACC_Testing"

print("Test Configuration:", testConfigName)

# Ensure CANoe module
canoeModul = gencache.EnsureModule(
    '{7F31DEB0-5BCC-11D3-8562-00105A3E017B}',
    0x0, 1, 54
)

# Find all .vtuexe files
testUnitPaths = [
    os.path.join(testUnitsFolder, file)
    for file in os.listdir(testUnitsFolder)
    if file.lower().endswith(".vtuexe")
]

if not testUnitPaths:
    raise FileNotFoundError(
        "No .vtuexe test units found in: " + testUnitsFolder
    )

# Add test units to CANoe
print("Trying to add test cases to CANoe...")

testConfigurations = CANoe.Configuration.TestConfigurations
testConfiguration = testConfigurations.Add()

testConfiguration.Name = testConfigName

testUnits = canoeModul.ITestUnits2(testConfiguration.TestUnits)

for testUnitPath in testUnitPaths:
    print("Adding TestUnit:", testUnitPath)
    testUnits.Add(testUnitPath)

# Start measurement
time.sleep(1)
CANoe.Measurement.Start()
print("Measurement started.")

time.sleep(2)

# Start test configuration
testConfiguration.Start()
print("Test configuration started.")

# Temporary wait - this does not independently confirm test completion
time.sleep(10)

print("Test execution wait period completed.")

# Stop measurement
try:
    if CANoe.Measurement.Running:
        CANoe.Measurement.Stop()
        print("Measurement stopped successfully.")
    else:
        print("Measurement is already stopped.")

except Exception as e:
    print("Measurement stop warning:", e)

# Open test report
print("Opening the vTEST report...")

report_file_path = (
    r"F:\Jenkins\Jenkins\Jenkins\ACC_RBS"
    r"\Report_AutomatedTests_VT_02.vtestreport"
)

if os.path.exists(report_file_path):
    print("Opening report:", report_file_path)
    os.startfile(report_file_path)
else:
    print("Report file not found:", report_file_path)
