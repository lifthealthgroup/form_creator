# Medical Assessment PDF Generator

The **Medical Assessment PDF Generator** is a web application that simplifies the process of generating PDF documents from medical assessment data stored in Excel files. This tool is designed for healthcare professionals to efficiently create and download necessary forms for patient assessments.

## Features

- **Multi-file Upload**: Supports simultaneous upload of multiple Excel files.
- **PDF Generation**: Automatically fills out PDF forms based on data extracted from the uploaded Excel files.
- **Highlighting Capabilities**: Highlights specific areas within the PDF to emphasize important information.
- **Zip File Download**: Downloads all generated PDFs as a single zip file for ease of access.

## Supported Forms

The following forms are included in the application:

- World Health Organization Disability Assessment Schedule (WHODAS)
- Care and Needs Scale (CANS)
- Life Skills Profile (LSP)
- Lawton Instrumental Activities of Daily Living (LAWTON)
- Berg Balance Scale (BBS)
- Lower Extremity Functional Scale (LEFS)
- Falls Risk Assessment Tool (FRAT)
- Health of the Nation Outcome Scales (HONOS)

## Technologies Used

- **Flask**: A lightweight web framework for Python that serves as the backbone of the application.
- **Pandas**: Utilized for data manipulation and analysis, particularly with Excel files.
- **PyMuPDF (fitz)**: Employed for PDF generation and manipulation.

## Getting Started

### Prerequisites

Ensure you have the following installed:

- Python 3.6+
- pip (Python package installer)

### Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/your-repo.git
   cd your-repo
   
2. **Set up a virtual environment (optional but recommended):**
   ```python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate
   
3. **Install the required packages:**
   ```
   pip install -r requirements.txt

### Usage
1. **Run the Flask application:**
   ```
   python app.py
2. **Access the application:** Open your web browser and navigate to http://127.0.0.1:5000.
3. **Upload your Excel files:** Use the provided interface to upload multiple medical assessment forms in Excel format.
4. **Download the generated PDFs:** After processing, a zip file containing all generated PDFs will be available for download.

## Contact
For questions or support, please contact [frank.snelling03@gmail.com].

## Acknowledgments
Special thanks to the open-source community for providing invaluable libraries and tools that made this project possible.
