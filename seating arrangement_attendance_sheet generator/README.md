📘 Exam Seating & Attendance Generator

Automated system for generating exam seating arrangements, subject-wise room allocation sheets, and attendance PDFs with student photos — all from a single Excel input file.

Built using Python, Pandas, Streamlit, ReportLab, and fully Dockerized for easy deployment.

🚀 Features

Single Excel Input
Reads timetable, course–roll mapping, roll–name mapping, and room capacities.

Automatic Clash Detection
Ensures no student is assigned to two exams in the same timeslot.

Smart Seating Allocation
Uses a greedy algorithm that:

Sorts subjects by descending strength

Allocates students room-wise based on effective capacity

Applies buffer seats and density mode (Dense / Sparse)

Attendance PDF Generation
Creates room-wise PDFs with:

Student photo

Name & roll number

Signature field

Exam/session metadata

Organized Output
Generates:

Room-wise and subject-wise Excel files

Attendance PDFs

Log files and error reports

Everything packaged into a downloadable ZIP

Streamlit Web Interface
Upload input file → Click → Download ZIP.

Docker Support
Run anywhere using a single command.

▶️ Running the App (Local)
pip install -r requirements.txt
streamlit run streamlit_app.py


🐳 Running via Docker
docker build -t seating-app .
docker run -p 8501:8501 seating-app

👉 http://localhost:8501