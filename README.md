# SEO-Description-Optimizer-for-Excel-Data
A Tkinter GUI app to optimize Excel/CSV descriptions for SEO using Open AI's API. Upload a file with "Name," "Id," "Description" columns, define an SEO prompt in the first row, and download an optimized Excel file. Features a vibrant interface with dark theme, gold accents, and teal/coral buttons. Ideal for enhancing content visibility.
## Features
- **Intuitive GUI**: User-friendly Tkinter interface with file upload/download dialogs.
- **File Support**: Processes `.xlsx` and `.csv` files with required columns ("Name," "Id," "Description").
- **SEO Optimization**: Enhances descriptions using Open AI’s `gpt-3.5-turbo` with a user-defined prompt.
- **Real-Time Feedback**: Displays status updates (e.g., “Processing…”, “Processing complete!”) in the GUI.
- **Error Handling**: User-friendly pop-up messages for invalid files or API issues.
- **Visual Design**: Dark blue-gray theme (`#1E1E2F`), gold text (`#FFD700`), teal upload button (`#00A896`), coral download button (`#FF6B6B`).
- **Output**: Saves optimized data to `seo_optimized_data.xlsx`.

## Prerequisites
- **Python**: 3.8+ (macOS).
- **Dependencies**: `pandas==2.2.3`, `openpyxl==3.1.5`, `openai==1.51.0`, `tkinter` (included with Python).
- **Open AI API Key**: Obtain from [Open AI](https://platform.openai.com).
