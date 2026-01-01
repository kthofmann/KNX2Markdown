# Project Architecture: KNX2Markdown

## Overview
**KNX2Markdown** is a Python CLI tool created to convert KNX Project exports (`.knxproj`) into human-readable Markdown documentation. It is designed for offline use.

## Core Implementation
The script `knx2markdown.py` operates in a linear pipeline:

### 1. Extraction
Unzips the `.knxproj` (which is a generic zip archive) in memory using Python's `zipfile` module. It specifically looks for `0.xml` (project topology) and `Hardware.xml` (product catalog).

### 2. Catalog Parsing
Before parsing the project devices, we build lookup tables to ensure human-readable output:
*   **Product Names**: `Hardware.xml` maps `ProductRefId` (e.g., `M-0083_H-...`) to text (e.g., "MDT Glass Push Button").
*   **Application Programs**: `M-*.xml` files map cryptic ComObject and Parameter IDs to their readable Names and Types (e.g., "Output 1 Switching" or "Cycle Time").

### 3. Module Resolution (Key Logic)
Many modern KNX devices (MDT, ABB, etc.) use a "Modular" structure where Object IDs are calculated dynamically based on a "Base Number".
*   **The Challenge**: `0.xml` only refers to `RefId="O-11"` but the Application Program defines `O-1` with a `ObjNumberBase`.
*   **The Solution**: The function `resolve_module_ref` implements the logic to match a specific "Module Instance" in the topology to its "Module Definition" and calculate the correct offset.

### 4. Structure Parsing
*   **Locations**: Uses `parse_locations` to traverse `knx:Locations` -> `knx:Space` (Building/Floor/Room).
*   **Topology**: Uses `parse_devices` to traverse `knx:Topology` -> `knx:Area` -> `knx:Line` -> `knx:DeviceInstance`.

### 5. Markdown Generation
The final report combines these data streams:
1.  **Project Stats**: Count of GA and Devices.
2.  **Visual Structure**: The physical locations (Room based) populated with the technical devices found in that room.
3.  **Device Details**: A table for each device showing its Parameters and Communication Objects (with Flags `C` `R` `W` `T` `U`).
4.  **Group Addresses**: A global list of all GAs.

## Configuration
*   **Language**: Controlled by `STRINGS` dict. Defaults to `de` (German). Switchable via `--lang`.
*   **Namespace**: Currently hardcoded to `http://knx.org/xml/project/23` (ETS 6).

## Data Flow Diagram
```mermaid
graph TD
    A[User] -->|Run script| B[knx2markdown.py]
    B -->|Unzip| C[.knxproj]
    C -->|Extract| D[0.xml (Topology)]
    C -->|Extract| E[Hardware.xml (Catalog)]
    C -->|Extract| F[M-x.xml (AppPrograms)]
    
    E -->|Build| G[Product Lookup]
    F -->|Build| H[Param/Obj Lookup]
    
    D -->|Parse| I[Locations & Devices]
    G --> I
    H --> I
    
    I -->|Generate| J[.md Report]
```
