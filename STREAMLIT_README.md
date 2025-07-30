# Lysna Assessment Analysis Tool - Streamlit Version

This is a web-based version of the Lysna Team Assessment Analysis Tool, converted from the original tkinter GUI application to a modern Streamlit web interface.

## Features

- **Web-based Interface**: Modern, user-friendly web interface accessible through any browser
- **File Upload**: Drag-and-drop or browse to upload CSV/Excel assessment files
- **Multi-language Support**: English, Danish, and Dutch language options
- **Real-time Progress**: Visual progress indicators during analysis
- **Data Preview**: Preview uploaded data before running analysis
- **Plot Generation**: Optional generation of visualisation charts
- **Dynamic Plots**: Enhanced plot colouring options
- **Downloadable Results**: All results packaged in a convenient ZIP file for download

## How to Use

1. **Start the Application**:
   
   **Option 1 - Command Line (All Platforms):**
   ```bash
   streamlit run streamlit_app.py
   ```
   
   **Option 2 - Windows Users:**
   ```bash
   run_streamlit_app.bat
   ```
   Simply double-click the `run_streamlit_app.bat` file or run it from the command prompt.

2. **Upload Your Data**:
   - Click "Browse files" or drag-and-drop your CSV or Excel file
   - Supported formats: `.csv`, `.xlsx`

3. **Configure Analysis**:
   - Select your preferred language (English, Danish, or Dutch)
   - Choose whether to generate plots
   - Enable dynamic plots for enhanced visualisation

4. **Run Analysis**:
   - Click "Start Analysis" to begin processing
   - Monitor progress with the log and progress indicator

5. **Download Results**:
   - Once complete, click "Download Results (ZIP)" to get all generated files

## What's Generated

The analysis produces:
- **Assessment results text files**: Detailed statistics and findings
- **Excel files**: Processed data and calculations
- **Attachment plots**: Visualisation of team attachment styles
- **Scale plots**: Various charts showing team performance metrics (if plots enabled)
- **Item plots**: A score distribution plot for each item
- **Variance analysis**: Statistical variance plots for different scales

## Installation

1. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Make sure you have the `lysna` package modules in the correct location

## Differences from .exe version

- **No file system browsing**: Upload files directly through the web interface
- **No output directory selection**: Results are automatically packaged for download
- **Real-time feedback**: Progress bars and status updates during processing
- **Data preview**: See a preview of your data before analysis
- **Responsive design**: Works on different screen sizes and devices
- **Cross-platform**: Runs by executing Python scripts to run app and analyses (requires more technical dependencies)

## Technical Notes

- Uses Streamlit for the web interface
- Maintains all original analysis functionality from the tkinter version
- Temporary directories are used for processing and cleaned up automatically
- Results are packaged as ZIP folder for easy download
- Supports the same file formats and analysis options as the original

## Troubleshooting

- **File upload issues**: Ensure your file is in CSV or Excel format and contains the required columns
- **Analysis errors**: Check that your data contains the expected assessment structure
- **Language errors**: Check that you have selected the correct analysis language for your data
- **Missing plots**: Ensure matplotlib and other plotting dependencies are installed

## Support

For issues or questions, refer to the original documentation or contact Frederik at frederik@bredgaard.dk.
