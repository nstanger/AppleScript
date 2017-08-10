tell application "Vivaldi" to set appName to name
tell script "Business Objects utilities" to init(appName)
tell script "Run EXRPT11" to run
