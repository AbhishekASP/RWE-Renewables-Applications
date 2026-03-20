# RDSPP Generator – Wind, PV & BESS Automation Tool

Overview

This project is a **desktop automation tool** developed to generate **RDSPP (Reference Designation System for Power Plants)** structures for renewable energy assets, including:

* 🌬️ Wind Turbines
* ☀️ Solar (PV) Systems
* 🔋 Battery Energy Storage Systems (BESS)

The tool automates the creation of **SAP Functional Locations and Equipment Lists**, significantly reducing manual effort and ensuring consistency across large-scale energy projects.

---

## Purpose

In large renewable energy projects, defining RDSPP structures manually is:

* Time-consuming
* Error-prone
* Difficult to scale across multiple assets (e.g., 50+ turbines)

This tool solves that by:

 Automating RDSPP hierarchy generation
 Standardizing SAP upload templates
 Supporting scalable replication across multiple units (G001–Gxxx)

---

##  Industry Context

This tool is inspired by and reflects **real-world workflows in renewable energy asset management**, particularly:

* SAP PM (Plant Maintenance) data structuring
* Functional Location hierarchy design
* Equipment master data creation

It demonstrates practical experience aligned with **energy companies like RWE Renewables**.

---

##  Key Features

###  1. Template-Based RDSPP Generation

* Upload engineering RDSPP templates
* Automatically detect structure (Code / RDSPP Code)
* Supports multiple template formats

---

### 2. Wind Turbine Automation

* Replicates turbine structures:

  ```
  G001 → G002 → G003 → ...
  ```
* Automatically generates:

  * Rotor blade systems (MDA11 → MDA12 → MDA13)
  * Yaw drive systems (MZ010 → MZ020 → MZ030)
* Preserves correct hierarchy and dependencies

---

###  3. PV & BESS System Support

* Handles solar and battery system structures
* Flexible hierarchy:

  ```
  F0 → F1 → F2 → P1
  ```
* Supports custom system configurations

---

###  4. Smart Data Transformation

* Converts engineering templates into SAP-ready format
* Automatically maps:

  * RDSPP Code → Functional Location
  * Code Level → Hierarchy Level
  * Functional Location Description → SAP Description

---

###  5. SAP Upload File Generation

Creates structured Excel outputs:

####  Functional Location Sheet

* Hierarchical structure (F0, F1)
* Planning & Maintenance plant assignment
* Validity dates and metadata

####  Equipment Sheet

* Equipment (F2, P1)
* Manufacturer, serial number, object type
* Linked to Functional Locations

---

###  6. GUI Application (PySide6)

* User-friendly desktop interface
* Load existing templates or generate new ones
* Dynamic input fields (turbines, serial numbers, etc.)
* One-click Excel export

---

##  Technical Highlights

* **Python (Core Logic)**
* **Pandas** for data processing
* **OpenPyXL** for Excel generation
* **PySide6** for GUI development

### Key Capabilities:

* Dynamic hierarchy generation
* Pattern-based code transformation
* Robust template parsing (multi-format support)
* Data validation and cleaning


##  Impact

This tool demonstrates:

* Real-world **data engineering in energy systems**
* Practical **automation of enterprise workflows**
* Strong understanding of:

  * SAP data structures
  * Renewable asset modeling
  * Scalable system design

---

##  Project Structure

```text
├── wind_generator.py        # Wind RDSPP generator
├── pv_bess_generator.py    # PV & BESS generator
├── gui_app.py              # PySide6 application
├── templates/              # Sample RDSPP templates
├── output/                 # Generated SAP files
```

---

##  Future Improvements

* YAML-based dynamic system configuration
* Support for additional renewable systems
* Integration with SAP APIs
* Enhanced UI/UX and validation

---

##  Author

**Abhishek Pillai**
M.Sc. Computer Science (Intelligent Systems)
Working Student – Data Management (Renewable Energy)


> Designed for efficiency. Built for scale. Applied in renewable energy systems.
