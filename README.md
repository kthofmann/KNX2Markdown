# KNX2Markdown

A Python tool that creates comprehensive reports from ETS project files (`.knxproj`) and exports them as Markdown. Ideal for documentation, GitHub Wikis, or Obsidian.

## Features

*   **Offline First**: Works completely locally, no cloud required.
*   **Detailed Parameters**: Extracts not only group addresses but also **device parameters** (e.g., delay times, brightness thresholds).
*   **Structured Overview**: Visualizes the building structure (Building -> Floor -> Room -> Devices).
*   **Markdown Output**: Generates clean Markdown with tables, icons, and status flags.

## Installation

Only Python 3 is required. No further dependencies.

1.  Download the script:
    ```bash
    curl -O https://raw.githubusercontent.com/kthofmann/KNX2Markdown/main/knx2markdown.py
    ```

## Usage

1.  Copy your `.knxproj` file into the same folder.
2.  Run:
    ```bash
    python3 knx2markdown.py
    ```
    The script automatically detects the `.knxproj` file and creates a matching `.md` file (e.g., `MyHouse_Export.md`).

Alternatively, you can specify the path directly:
```bash
python3 knx2markdown.py /path/to/your/project.knxproj
```

## Language Options

The output language defaults to **German** (`de`).
You can switch to English using the `--lang` argument:

```bash
python3 knx2markdown.py --lang en
```

Supported languages:
*   `de`: German (Default)
*   `en`: English

## Example Output

The generated report includes:
*   Project statistics
*   Building structure with devices
*   Group Addresses table
*   Detailed device list with flags and parameters

## License

MIT License
