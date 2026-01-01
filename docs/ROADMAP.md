# Project Roadmap & History

## Active Tasks
- [ ] Future: Support for older ETS namespaces (ET S5/4)
- [ ] Future: Web-based viewer option

## Completed History
- [x] Initial Script Implementation
    - Basic parsing of `0.xml`
    - ZIP extraction logic
- [x] Feature: Auto-detection of KNXProj
    - Scan current directory for `.knxproj`
    - Remove hardcoded paths ('HofmannsHome')
- [x] Feature: Human-readable Product IDs
    - Parse `Hardware.xml` for product names
    - Map `ProductRefId` to text
- [x] Fix: Device Presence Detector display issues
    - Ensure consistent formatting for all device types
- [x] Feature: Reorder Report Structure
    - Order: Building Structure -> Devices -> Parameters -> Connections -> Addresses
- [x] Documentation
    - Created `docs/ARCHITECTURE.md`
    - Created `docs/ROADMAP.md`
