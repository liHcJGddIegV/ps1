# Define an array of new folder names
$folders = @(
    "Calibration Docs",
    "Data Saves - Program - Wiring Diagram",
    "H-Frames",
    "Instrumentation",
    "Civil test reports and certs",
    "Commissioning Form",
    "Commissioning documents",
    "P&P Bond",
    "Stamped Prints",
    "Submittal"
)

# Loop through each folder name in the array
foreach ($folder in $folders) {
    # Create a new directory with the name from the array
    New-Item -Path $folder -ItemType Directory
}
