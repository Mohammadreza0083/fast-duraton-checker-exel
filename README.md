# 🎥 Fast Duration Checker Excel

A simple Python tool to scan video course folders, extract video durations using ffprobe, and generate an interactive Excel workbook to track your course progress.

## ✨ Features
- 🔍 Scans all video files (.mp4, .mkv, .mov, .avi, .webm) inside a folder and its subfolders
- 📁 Groups videos by their folder (considered as Section) and lists each video as a Subsection
- ⏱️ Extracts video durations automatically
- 📊 Generates an Excel file with:
  - ✅ A checklist column (0/1) to mark watched videos
  - 🎨 Conditional formatting that colors completed sections green and incomplete sections red
  - 📈 Calculates and displays:
    - Progress percentage per section
    - Time watched and remaining per section (in minutes and formatted hours:minutes)
    - Total course progress and time remaining
  - 🔒 Data validation for the watched column to allow only 0 or 1
  - 📐 Well formatted with bold headers, borders, and column width adjustments

## 🛠️ Installation
Make sure you have these installed:
- Python 3.x
- ffprobe (part of FFmpeg) — must be accessible from your system PATH
- Python packages:
```bash
pip install openpyxl
```

## 🚀 Usage
1. Place your course folder (with subfolders as sections and videos inside) in any directory
2. Run the script inside the root folder or specify the path in the code
3. The script will generate an Excel file named `course_progress.xlsx` with all the details
4. Open the Excel file and mark videos you have watched by changing 0 to 1 in the "Watched (0/1)" column
5. The progress and time calculations update automatically

## ⚙️ How it works
- Uses ffprobe to get the duration of each video file
- Collects data in tuples of (section, video, duration)
- Creates an Excel sheet listing all sections and subsections with durations
- Adds formulas to compute progress percentages and time summaries
- Applies conditional formatting to visually distinguish completed and incomplete sections

## 📁 Example folder structure
```
CourseRootFolder/
├── Section 1/
│   ├── lesson1.mp4
│   ├── lesson2.mp4
├── Section 2/
│   ├── part1.mkv
│   ├── part2.mkv
```

## 📋 Requirements
- Python 3.x
- FFmpeg with ffprobe
- Python packages: openpyxl

## 👨‍💻 Developer
- **Mohammadreza Bonyadi**

## 📄 License
MIT License

---
Made with ❤️ by Mohammadreza Bonyadi

