# Lysna automated analysis
This repo contains code and materials (cleaned of proprietary information) for the automated analysis of the survey-based [Lysna](https://lysna.world/) team assessment.

## The Lysna assessment
The Lysna team assessment is a survey-based assessment of team dynamics, including psychological safety, attachment-informed characteristics, and other group metrics relevant to team culture and performance.

## Why automate it?
In the past, once a team completed the survey, a Lysna consultant would manually analyse the data to extracting key findings from the results, often having to spend several days of work to go in-depth on each subscale of the assessment.
This application solves that problem by automating the preprocessing and statistical analysis of the assessment.

This results in
1. time-savings;
2. standardising the delivered product; and,
3. produces visuals to support the reporting and delivery of key findings.


# Lysna Team Assessment Analysis Tool

The Lysna Team Assessment Analysis Tool is a comprehensive solution for analysing team dynamics and attachment styles. This repository contains both a modern web-based Streamlit application and traditional executable versions for different platforms.

## Two Ways to Use the Tool

### Option 1: Web Application (Recommended)
**Modern Streamlit-based web interface** - Perfect for users comfortable with Python environments.

**Features:**
- **Web-based Interface**: Modern, user-friendly web interface accessible through any browser
- **File Upload**: Drag-and-drop or browse to upload CSV/Excel assessment files  
- **Multi-language Support**: English, Danish, and Dutch language options
- **Real-time Progress**: Visual progress indicators during analysis
- **Data Preview**: Preview uploaded data before running analysis
- **Plot Generation**: Optional generation of visualisation charts
- **Dynamic Plots**: Enhanced plot colouring options
- **Downloadable Results**: All results packaged in a convenient ZIP file for download

**How to Start:**
- **Command Line (All Platforms):** `streamlit run streamlit_app.py`
- **Windows Users:** Run `run_streamlit_app.bat` or double-click the batch file

**Requirements:** Python environment with dependencies from `requirements.txt`

### Option 2: Executable Applications
**Ready-to-run .exe files** - Perfect for users who prefer standalone applications.

**Available Versions:**
- **Windows**: `app.exe` (main version)
- **macOS**: Available through GitHub Actions artifacts
- **Language-specific versions**: Found in Legacy and Proprietary folders

**Benefits:**
- No Python installation required
- Traditional GUI interface
- Direct file system access
- Standalone operation

## 📋 What the Analysis Produces

Both versions generate identical analysis outputs:
- **Assessment results text files**: Detailed statistics and findings
- **Excel files**: Processed data and calculations  
- **Attachment plots**: Visualisation of team attachment styles
- **Scale plots**: Various charts showing team performance metrics
- **Item plots**: Score distribution plots for each assessment item
- **Variance analysis**: Statistical variance plots for different scales

## 🚀 Quick Start

### For Web Application:
1. Install dependencies: `pip install -r requirements.txt`
2. Run: `streamlit run streamlit_app.py` or use `run_streamlit_app.bat` (Windows)
3. Upload your CSV/Excel file through the web interface
4. Configure language and plot options
5. Download results as ZIP file

### For Executable Application:
1. Download the executable (see below for Windows and macOS instructions)
2. Run the executable
3. Select your data file and desired output folder
4. Configure language and plot options
5. Run

## 📥 Download Instructions

### Windows Executable Download:

#### Step 1: Find your file
You most likely need the **app.exe**.

Click the name of the desired file (e.g. what's highlighted in blue below). Make sure to check the file extension (look for files that end in .exe).

![filename](filename.png)

#### Step 2: Download
Click the download button in the top right corner.

![download_button](step2.png)

#### Step 3: If prompted, proceed to download anyway
Microsoft may warn you about potentially harmful programmes as they're not familiar with Lysna as a developer. If this happens, simply proceed to download anyway - this may require clicking an "advanced settings" button or similar.

### macOS Application Download:

#### Step 1: Navigate to the Actions Tab
Click on the Actions tab at the top of the page.

#### Step 2: Find the Workflow Run
Locate the most recent successful workflow run with the name "Build macOS application".
![workflow_name](workflow_name.png)

Click on the workflow run to view its details.

#### Step 3: Download the Artifact
Scroll down to the Artifacts section.
You should see an artifact named lysna-macos-app.
Click on the artifact name to download it.

## 🛠 Technical Details

### Web Application
- **Framework**: Streamlit
- **Languages**: Python
- **Dependencies**: Listed in `requirements.txt`
- **Output**: ZIP download with all results
- **Cross-platform**: Works on Windows, macOS, Linux

### Executable Applications  
- **Build Tool**: PyInstaller
- **Languages**: Multi-language support (EN, DA, NL)
- **Output**: Direct file system output
- **Platform-specific**: Separate builds for Windows/macOS

## 📁 Repository Structure

```
├── streamlit_app.py          # Web application main file
├── run_streamlit_app.bat     # Windows batch launcher
├── STREAMLIT_README.md       # Detailed web app documentation
├── app.exe                   # Main Windows executable
├── requirements.txt          # Python dependencies
├── lysna/                    # Core analysis modules
├── Legacy versions/          # Previous executable versions
├── Proprietary versions/     # Specialised versions
└── Results/                  # Sample analysis outputs
```

## 🤝 Support

For issues or questions:
- **Web Application**: See `STREAMLIT_README.md` for detailed documentation
- **General Support**: Contact Frederik at frederik@bredgaard.dk
- **Technical Issues**: Check the troubleshooting sections in the documentation

## 🔄 Version Comparison

| Feature | Web Application | Executable |
|---------|----------------|------------|
| **Installation** | Requires Python | Standalone |
| **Interface** | Modern web UI | Traditional GUI |
| **File Handling** | Upload/Download | Direct filesystem |
| **Progress Tracking** | Real-time web progress | Traditional progress bars |
| **Platform Support** | Cross-platform | Platform-specific builds |
| **Updates** | Pull latest code | Download new executable |
| **Data Preview** | Built-in preview | File-based |

Choose the option that best fits your technical comfort level and usage requirements!
