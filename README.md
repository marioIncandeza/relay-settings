# SEL Settings Generator

This tool automates the generation of Schweitzer Engineering Laboratories (SEL) relay settings files (RDB format) by merging configuration data from an Excel spreadsheet with a set of template text files.

## Features

*   **Excel Integration**: Reads settings values directly from structured Excel tables.
*   **Template Based**: Uses a folder of template `.txt` files to generate the final RDB configuration structure.
*   **Multi-Relay Support**: Configured to support various SEL devices including:
    *   Feeder (351S)
    *   HV (351S)
    *   Transformer (487E, 787)
    *   Capacitor Bank (487V)
    *   Bus Differential (587Z)
    *   Meter (735)
    *   Automation Controller (DPAC 2440)
    *   Line Protection (411L)
*   **Web Interface**: Modern GUI built with Streamlit.
*   **Legacy GUI**: Includes a Tkinter-based desktop interface.

## Installation

1.  **Clone the repository** to your local machine.
2.  **Set up a Virtual Environment** (Recommended):
    ```bash
    python -m venv venv
    .\venv\Scripts\Activate  # Windows
    # source venv/bin/activate  # Mac/Linux
    ```
3.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

### Web Application (Streamlit)
This is the recommended interface.

1.  Run the application:
    ```bash
    streamlit run streamlit_app.py
    ```
2.  The app will open in your default web browser.
3.  **Select Relay Type**: Choose the device you are configuring from the sidebar.
4.  **Upload Config**: Upload your settings Excel file (`.xlsx`).
5.  **Choose Template**:
    *   **Embedded**: Uses the templates stored in the `templates/` directory.
    *   **Custom**: Upload a `.zip` file containing your specific template files.
6.  **Generate**: Click "Generate Settings" to create and download a `.zip` file containing your configured RDB files.

### Desktop Application (Legacy)
1.  Run the Tkinter app:
    ```bash
    python main.py
    ```

## Project Structure

*   `streamlit_app.py`: Main entry point for the Web GUI.
*   `rdb.py`: Core logic for parsing Excel and processing RDB text files. Contains `gen_settings` and `get_template_info`.
*   `main.py`: Legacy Tkinter GUI entry point.
*   `templates/`: Directory containing default templates for each relay type.
*   `requirements.txt`: Python dependencies.

## Template Info
The application reads metadata from `Misc/Cfg.txt` within each template folder (if available) to display relevant information (Part Number, Firmware ID) in the sidebar.