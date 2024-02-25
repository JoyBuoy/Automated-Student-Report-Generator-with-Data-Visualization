# Automated-Student-Report-Generator-with-Data-Visualization

This Python script is your report card hero! It generates personalized reports for each student in your Excel file, complete with final grades, class stats, and bar graphs. No more manual calculations or boring reports - automate your grading process with this easy-to-use tool.


<h1>Description:</h1>
Tired of manually creating student reports? This Python script automates the process, generating personalized PDF reports for each student with:

    Final Grades: Based on weighted averages of test scores.
    Class Statistics: Average scores and highest marks in each subject.
    Bar Graphs: Comparing student marks to class averages.
    Student Pictures: Student photos for added personalization.

<h2>Key Features:</h2>

    Easy to Use: Simply provide an Excel file with student data.
    Visually Engaging: Includes bar graphs for easy comparison.
    Customizable: Modify weights for final grade calculation and paths for files and folders.
    Handles Missing Data: Gracefully handles students with incomplete information.

<h2>Requirements:</h2>

    Python 3.x
    pandas
    matplotlib
    reportlab

<h2>Usage:</h2>

    Install required libraries: pip install pandas matplotlib reportlab
    Update paths in the script:
        excel_file_path: Path to your Excel file with student data.
        output_folder_plot: Folder to store generated graphs.
        student_photos : Folder containing student photos.
        report_folder : Folder to store generated final reports.
    Run the script: python script.py

<h2>Generated Reports:</h2>

    Individual PDF reports for each student will be created in the specified output folder.
    Each report includes:
        Student ID (and name, if available)
        Photo 
        Final grade
        Individual subject marks
        Class statistics
        Bar graph comparing student marks to class averages

<h2>Contribute:</h2>
Feel free to fork the repository and suggest improvements!

<h2>Note:</h2>
This readme provides a general overview. For detailed information, refer to the script itself.
In addition make sure to name the student images with their respective student ID's.
