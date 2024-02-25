import os
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas


excel_file_path = 'C:\\Users\\User\\Music\\student marks\\Book2.xlsx'
output_folder_plot = 'C:\\Users\\User\\Music\\student marks\\student_report_plot'
student_photos = 'C:\\Users\\User\\Music\\student marks\\student_photos'
report_folder = 'C:\\Users\\User\\Music\\student marks\\student_reports'

test_1_weightage = 0.3
test_2_weightage = 0.3
test_3_weightage = 0.4


# It returns a final grade.
#calculate_final_grade to assign grades.

def calculate_final_grade(test_1, test_2, test_3):

    # Calculate the final grade based on the specified weights
    final_grade = test_1_weightage * test_1 + test_2_weightage * test_2 + test_3_weightage * test_3

    # Assign letter grades based on the final grade range
    if final_grade >= 90:
        return 'A+'
    elif 75 <= final_grade < 90:
        return 'A'
    elif 60 <= final_grade < 75:
        return 'B'
    elif 40 <= final_grade < 60:
        return 'C'
    else:
        return 'D (Needs Improvement)'


def generate_student_report(data_frame, student_id, report_folder):

    # Filter data for the given student_id
    student_data = data_frame[data_frame['Student_ID'] == student_id].copy()

    # Check if student_data is not empty
    if student_data.empty:
        print(f"No data found for Student ID: {student_id}")
        return

    # Calculate final grade for each subject
    student_data['Final_Grade'] = student_data.apply(
        lambda row: calculate_final_grade(row['test-1'], row['test-2'], row['test-3']), axis=1)

    # Create the output directory if it doesn't exist
    os.makedirs(report_folder, exist_ok=True)

    # Create a PDF report
    report_pdf_path = os.path.join(report_folder, f'student_{student_id}_report.pdf')
    c = canvas.Canvas(report_pdf_path)

    # Set starting point for drawing text
    y_position = 750

    # Add new page if necessary.
    def add_new_page():
        nonlocal y_position
        c.showPage()  # Start a new page
        y_position = 750  # Reset Y position

    # Display student picture
    image_path = os.path.join(student_photos, f'{student_id}.jpg')
    if os.path.exists(image_path):
        c.drawInlineImage(image_path, 400, 650, width=100, height=100)

    c.drawString(100, y_position, f'Student ID: {student_id}')
    y_position -= 20

    # Checking if the "Name" column exists in student_data
    if "Name" in student_data.columns:
        c.drawString(100, y_position, f'Name: {student_data["Name"].iloc[0]}')
    else:
        c.drawString(100, y_position, 'Name: Not Available')

    # Calculate class statistics
    class_size = len(data_frame)
    class_average = data_frame[['test-1', 'test-2', 'test-3']].mean().mean()
    highest_in_class = data_frame[['test-1', 'test-2', 'test-3']].max().max()

    y_position -= 20
    c.drawString(100, y_position, f'Class Size: {class_size}')
    y_position -= 20
    c.drawString(100, y_position, f'Class Average: {class_average:.2f}')
    y_position -= 20
    c.drawString(100, y_position, f'Highest in Class: {highest_in_class}')

    # Calculate and display the final grade
    final_grade = student_data['Final_Grade'].values[0]
    y_position -= 20
    c.drawString(100, y_position, f'Final Grade: {final_grade}')

    # Display individual subject marks
    subjects = ['test-1', 'test-2', 'test-3']
    for subject in subjects:
        y_position -= 20
        c.drawString(100, y_position, f'Student\'s Marks in {subject}: {student_data[subject].values[0]}')
        y_position -= 20
        c.drawString(100, y_position, f'Class Average in {subject}: {data_frame[subject].mean():.2f}')

    # Adding space before the bar graph.
    y_position -= 50
    # starting point for the bar graph
    y_position -= 20
    # mentioning the bar width
    bar_width = 0.2
    student_marks = [student_data[subject].values[0] for subject in subjects]
    class_averages = [data_frame[subject].mean() for subject in subjects]

    # Check if there is enough space for the graph, if not add a new page
    if y_position < 400:
        add_new_page()

    # Plotting the bar graph
    colors = ['blue', 'green', 'red']
    for i, (subject, student_mark, class_average) in enumerate(zip(subjects, student_marks, class_averages)):
        plt.bar(i, student_mark, color=colors[i], width=bar_width, label=f'Student Marks {subject}')
        plt.bar(i + bar_width, class_average, color=colors[i], width=bar_width, label=f'Class Average {subject}', alpha=0.5)

    plt.title('Marks Comparison')
    plt.xlabel('Subjects')
    plt.ylabel('Marks')
    plt.legend()

    # Save the plot to the output folder
    plot_image_path = os.path.join(output_folder_plot, f'student_{student_id}_marks_comparison.png')
    plt.savefig(plot_image_path)
    plt.close()

    # Draw the bar graph in the PDF
    y_position -= 350
    c.drawInlineImage(plot_image_path, 100, y_position, width=400, height=300)
    # Saving the PDF
    c.save()


# Read the Excel file using panda library
data_frame = pd.read_excel(excel_file_path)

# Example usage for generating the report for all students in the class
for student_id in data_frame['Student_ID']:
    generate_student_report(data_frame, student_id, report_folder)
