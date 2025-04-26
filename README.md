This Python script reads values from a values.txt file and writes them into an output.xlsx Excel file, one value per row.
It runs automatically at a set interval and handles missing files or locked Excel files gracefully.

Features:
• Automatically creates values.txt if it doesn't exist

• Parses numbers (integers and floats) correctly, leaves other text as-is

• Continuously runs at a customizable interval (default: every 1 minute)

• Handles locked Excel files (prompts user if the output file is open)

• Simple and lightweight, designed for easy automation
